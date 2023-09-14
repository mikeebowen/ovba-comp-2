// vba-comp.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include <iostream>
#include <fstream>
#include <vector>
#include <string>
#include <iomanip>
#include <sstream>

// Create struct to hold state
struct State
{
	int compressedRecordEnd;
	int compressedCurrent;
	int compressedChunkStart;
	int decompressedCurrent;
	int decompressedBufferEnd;
	int decompressedChunkStart;
};

struct CopyToken
{
	uint16_t offset;
	uint16_t length;
};

struct CopyTokenHelpData
{
	uint16_t lengthMask;
	uint16_t offsetMask;
	uint16_t bitCount;
	uint16_t maximumLength;
};

std::vector<uint8_t> *readFileIntoByteArray(const std::string &filename)
{
	std::ifstream file(filename, std::ios::binary);

	if (file.fail())
	{
		throw std::runtime_error("Failed to open file");
	}

	// Determine file size
	file.seekg(0, std::ios::end);
	const std::streamsize size = file.tellg();
	file.seekg(0, std::ios::beg);

	// Read file into byte array using heap allocation
	auto byteArray = new std::vector<uint8_t>(size);

	if (!file.read((char *)byteArray->data(), size))
	{
		delete byteArray;
		throw std::runtime_error("Failed to read file");
	}

	file.close();

	return byteArray;
}

std::string hexStr(unsigned char* data, int len)
{
	constexpr char hexmap[] = { '0', '1', '2', '3', '4', '5', '6', '7',
						   '8', '9', 'a', 'b', 'c', 'd', 'e', 'f' };

	std::string s(len * 2, ' ');

	for (int i = 0; i < len; ++i)
	{
		s[2 * i] = hexmap[(data[i] & 0xF0) >> 4];
		s[2 * i + 1] = hexmap[data[i] & 0x0F];
	}
	return s;
}

void writeBytesToFile(const std::string &filename, const std::vector<uint8_t> &bytes)
{
	// Open the file for writing in binary mode
	std::ofstream file(filename, std::ios::binary | std::ios::app);

	if (!file.is_open())
	{
		// File doesn't exist, so create a new one
		file.open(filename, std::ios::binary);
	}

	// Write the vector of bytes to the file
	file.write(reinterpret_cast<const char *>(bytes.data()), bytes.size());

	// Close the file
	file.close();
}

void writeHexToFile(const std::string &filename, const std::vector<uint8_t> &bytes)
{
	// Open the file for writing in text mode
	std::ofstream file(filename);

	// Write the hex values of the bytes to the file
	for (auto b : bytes)
	{
		std::string s = hexStr(&b, 1);
		file << "0x" << s << " ";
	}

	// Close the file
	file.close();
}
// 2.4.1.3.12 Extract CompressedChunkSize (https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/4994c768-d35d-497d-937d-d577611cb17f)
uint16_t extractCompressedChunkSize(uint16_t *header)
{
	// SET temp TO Header BITWISE AND 0x0FFF
	return (*header & 0x0FFF) + 3;
}

// 2.4.1.3.17 Extract FlagBit (https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/ff4680fd-b2af-4d3f-bb86-7d5d296218f60)
int extractFlagBit(unsigned int index, uint8_t flagByte)
{
	return (flagByte >> index) & 1;
}

// 2.4.1.3.15 Extract CompressedChunkFlag (https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/a954990f-b3d3-43e0-8d52-3d86d1f5b9af)
unsigned int extractCompressedChunkFlag(uint16_t *header)
{
	return (*header & 0x8000) >> 15;
}

// 2.4.1.3.19.1 CopyToken Help (https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/220bde4c-02b5-41ef-9f9a-8608718ce913)
CopyTokenHelpData copyTokenHelp(State *state, uint16_t copyToken)
{
	CopyTokenHelpData copyTokenHelpData;

	// SET difference TO DecompressedCurrent MINUS DecompressedChunkStart
	int difference = state->decompressedCurrent - state->decompressedChunkStart;
	// SET BitCount TO the smallest integer that is GREATER THAN OR EQUAL TO LOGARITHM base 2 of difference
	int bitCount = ceil(log2(difference));
	// SET BitCount TO the maximum of BitCount and 4
	copyTokenHelpData.bitCount = std::max(bitCount, 4);
	// SET LengthMask TO 0xFFFF RIGHT SHIFT BY BitCount
	copyTokenHelpData.lengthMask = 0xFFFF >> copyTokenHelpData.bitCount;
	// SET OffsetMask TO BITWISE NOT LengthMask
	copyTokenHelpData.offsetMask = ~copyTokenHelpData.lengthMask;
	// SET MaximumLength TO (0xFFFF RIGHT SHIFT BY BitCount) PLUS 3
	copyTokenHelpData.maximumLength = (0xFFFF >> copyTokenHelpData.bitCount) + 3;

	return copyTokenHelpData;
}

// 2.4.1.3.19.2 Unpack CopyToken (https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/6a77ef81-79da-41a9-9b7b-a4c807f326d5)
CopyToken unpackCopyToken(State *state, uint16_t token)
{
	CopyToken copyToken;

	CopyTokenHelpData copyTokenHelpData = copyTokenHelp(state, token);

	copyToken.length = (token & copyTokenHelpData.lengthMask) + 3;
	uint16_t temp1 = token & copyTokenHelpData.offsetMask;
	uint16_t temp2 = 16 - copyTokenHelpData.bitCount;
	copyToken.offset = (temp1 >> temp2) + 1;

	return copyToken;
}

// 2.4.1.3.11 Byte Copy(https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/7b75cf79-b736-47db-96ab-a443636518a8)
void byteCopy(std::vector<uint8_t> *decompressedBytes, State *state, int copySource, int byteCount)
{
	int srcCurrent = copySource;
	int dstCurrent = state->decompressedCurrent;

	for (int count = 1; count <= byteCount; count++)
	{
		decompressedBytes->insert(decompressedBytes->begin() + dstCurrent, (*decompressedBytes)[srcCurrent]);

		srcCurrent++;
		dstCurrent++;
	}
}

// 2.4.1.3.5 Decompressing a Token (https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/d7069c13-458a-4020-83b3-1f3f49450e9e)
void decompressToken(std::vector<uint8_t> *compressedBytes, std::vector<uint8_t> *decompressedBytes, State *state, unsigned int index, uint8_t flagByte)
{
	// CALL Extract FlagBit (section 2.4.1.3.17) with index and Byte returning Flag
	int flag = extractFlagBit(index, flagByte);

	if (flag == 0)
	{
		decompressedBytes->insert(decompressedBytes->begin() + state->decompressedCurrent, (*compressedBytes)[state->compressedCurrent]);

		state->decompressedCurrent++;
		state->compressedCurrent++;
	}
	else if (flag == 1)
	{
		uint16_t foo = (*compressedBytes)[state->compressedCurrent + 1];
		uint16_t bar = (*compressedBytes)[state->compressedCurrent];
		// SET Token TO the CopyToken (section 2.4.1.1.8) at CompressedCurrent
		uint16_t token = static_cast<uint16_t>((*compressedBytes)[state->compressedCurrent + 1]) << 8 | (*compressedBytes)[state->compressedCurrent];
		CopyToken copyToken = unpackCopyToken(state, token);

		// SET CopySource TO DecompressedCurrent MINUS Offset
		int copySource = state->decompressedCurrent - copyToken.offset;

		byteCopy(decompressedBytes, state, copySource, copyToken.length);

		state->decompressedCurrent += copyToken.length;
		state->compressedCurrent += 2;
	}
	else
	{
		throw std::runtime_error("Invalid flag");
	}
}

// 2.4.1.3.4 Decompressing a TokenSequence (https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/70ea7393-cc0d-4efd-ae18-853f29a062c80
void decompressTokenSequence(std::vector<uint8_t> *compressedBytes, std::vector<uint8_t> *decompressedBytes, State *state, int compressedEnd)
{
	// SET Byte TO the FlagByte (section 2.4.1.1.7) located at CompressedCurrent
	uint8_t flagByte = (*compressedBytes)[state->compressedCurrent];
	// INCREMENT CompressedCurrent
	state->compressedCurrent++;

	if (state->compressedCurrent < compressedEnd)
	{
		for (int i = 0; i <= 7; i++)
		{
			if (state->compressedCurrent < compressedEnd)
			{
				decompressToken(compressedBytes, decompressedBytes, state, i, flagByte);
			}
		}
	}
}

// 2.4.1.3.3 Decompressing a RawChunk (https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/86ff30e6-9442-4232-ba00-e616c64fd9ac0
void decompressRawChunk(std::vector<uint8_t> *compressedBytes, std::vector<uint8_t> *decompressedBytes, State *state)
{
	// APPEND 4096 bytes from CompressedCurrent TO DecompressedCurrent
	for (int i = 0; i < 4096; i++)
	{
		decompressedBytes->insert(decompressedBytes->begin() + state->decompressedCurrent, (*compressedBytes)[state->compressedCurrent]);

		state->compressedCurrent++;
		state->decompressedCurrent++;
	}
}

// 2.4.1.3.2 Decompressing a CompressedChunk (https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/3d5ea4df-e8a5-4079-a454-9595b840525f)
void decompressCompressedChunk(std::vector<uint8_t> *compressedBytes, std::vector<uint8_t> *decompressedBytes, State *state)
{
	// SET Header TO the CompressedChunkHeader (section 2.4.1.1.5) located at CompressedChunkStart
	uint16_t header = static_cast<uint16_t>((*compressedBytes)[state->compressedChunkStart + 1]) << 8 | (*compressedBytes)[state->compressedChunkStart];
	// CALL Extract CompressedChunkSize (section 2.4.1.3.12) with Header returning Size
	uint16_t compressedChunkSize = extractCompressedChunkSize(&header);
	// CALL Extract CompressedChunkFlag (section 2.4.1.3.15) with Header returning CompressedFlag
	unsigned int compressedChunkFlag = extractCompressedChunkFlag(&header);
	// SET DecompressedChunkStart TO DecompressedCurrent
	state->decompressedChunkStart = state->decompressedCurrent;
	// SET CompressedEnd TO the minimum of CompressedRecordEnd and (CompressedChunkStart PLUS Size)
	int compressedEnd = std::min(state->compressedRecordEnd, state->compressedChunkStart + compressedChunkSize);
	// SET CompressedCurrent TO CompressedChunkStart PLUS 2
	state->compressedCurrent = state->compressedChunkStart + 2;

	if (compressedChunkFlag == 1)
	{
		while (state->compressedCurrent < compressedEnd)
		{
			decompressTokenSequence(compressedBytes, decompressedBytes, state, compressedEnd);
		}
	}
	else
	{
		decompressRawChunk(compressedBytes, decompressedBytes, state);
	}
}

int main()
{
	// Read compressed bytes from file into vector and create vector for decompressed bytes
	std::vector<uint8_t> *compressedBytes = readFileIntoByteArray("C:\\source\\scripts\\rust\\ovba-comp\\compressed.bin");
	std::vector<uint8_t> *decompressedBytes = new std::vector<uint8_t>;

	// Set up state
	State *state = new State;

	state->compressedRecordEnd = compressedBytes->size();
	state->compressedCurrent = 0;
	state->compressedChunkStart = 0;
	state->decompressedCurrent = 0;
	state->decompressedBufferEnd = 0;
	state->decompressedChunkStart = 0;

	// 2.4.1.3.1 Decompression Algorithm (https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/492124cc-5afc-48c8-b439-b42ad7087a7b)
	if ((*compressedBytes)[state->compressedCurrent] == 0x01)
	{
		state->compressedCurrent++;

		while (state->compressedCurrent < state->compressedRecordEnd)
		{
			state->compressedChunkStart = state->compressedCurrent;

			decompressCompressedChunk(compressedBytes, decompressedBytes, state);
		}
	}
	else
	{
		throw std::runtime_error("Invalid compression flag");
	}

	// Write decompressed bytes to file
	writeBytesToFile("C:\\source\\scripts\\cpp\\vba-comp\\decompressed.bin", *decompressedBytes);
	writeHexToFile("C:\\source\\scripts\\cpp\\vba-comp\\decompressed.txt", *decompressedBytes);

	delete state;
	delete decompressedBytes;
	delete compressedBytes;
}
