/*
 * Author   : Berry de Groot
 * Date     : 02-11-2014
 * Document : [MS-OVBA] (http://msdn.microsoft.com/en-us/library/cc313094%28v=office.12%29.aspx)
 * Specific : Decompression Algorithm (http://msdn.microsoft.com/en-us/library/dd923471%28v=office.12%29.aspx) 
 *
 * Known (security) issues:
 * 
 * 1 - Not enough boundary checks on data input, suspected to reading outside of buffer. (data exposure)
 * 2 - Not enough boundary checks on created output, suspected to writing outside of buffer. (buffer overflow)
 * 3 - Not enough boundary checks on created output, suspected to reading infront of buffer and exposing data. (data exposure)
 *
 * Advice:
 * Do not use this with untrusted input without fixing above issues!
 *
 */


#include "stdafx.h"
#include "MS_OVBA_decompression.h"

#define DECOMPRESSED_BUFFER_SIZE(state) ((size_t)(state.dbs.decompressedCurrent - state.db))
#define MIN(a,b) ((a)<(b)?(a):(b))
#define MAX(a,b) ((a)>(b)?(a):(b))

typedef struct {
	unsigned int compressedChunkSize : 12;
	unsigned int compressedChunkSignature : 3;
	unsigned int compressedChunkFlag : 1;
} CompressedChunkHeader;

typedef struct {
	CompressedChunkHeader header;
	// unsigned char data[]; 
} CompressedChunk;

typedef struct {
	unsigned char *compressedRecordEnd;
	unsigned char *compressedCurrent;
} CompressedContainerState;

typedef struct {
	unsigned char *compressedChunkStart;
} CompressedChunkState;

typedef struct {
	unsigned char *decompressedCurrent;
	unsigned char *decompressedBufferEnd;
} DecompressedBufferState;

typedef struct {
	unsigned char *decompressedChunkStart;
} DecompressedChunkState;

typedef struct {
	CompressedContainerState ccs;
	CompressedChunkState cchs;
	DecompressedChunkState dchs;
	DecompressedBufferState dbs;
	unsigned char *db;
} MS_OVBA_state;

void byteCopy(void *dst, void *src, size_t size)
{
	unsigned char *d = (unsigned char *)dst;
	unsigned char *s = (unsigned char *)src;

	for (unsigned int i = 0; i < size; i++)
	{
		d[i] = s[i];
	}
}

void copyTokenHelp 
(
	MS_OVBA_state *state,
	unsigned __int16 *lengthMask,
	unsigned __int16 *offsetMask,
	unsigned __int16 *bitCount,
	unsigned __int16 *maxLength,
	unsigned __int16 *token
)
{
	size_t difference = state->dbs.decompressedCurrent - state->dchs.decompressedChunkStart;

	*bitCount = MAX((unsigned __int16)(ceil(log2((double)difference))), 4);

	*lengthMask = 0x0ffff >> *bitCount;
	*offsetMask = ~(*lengthMask);
	*maxLength = (0x0ffff >> *bitCount) + 3;
}

void unpackCopyToken(MS_OVBA_state *state, unsigned __int16 *offset, unsigned __int16 *length, unsigned __int16 *token)
{

	unsigned __int16 lengthMask;
	unsigned __int16 offsetMask;
	unsigned __int16 bitCount;
	unsigned __int16 maxLength;

	copyTokenHelp(state, &lengthMask, &offsetMask, &bitCount, &maxLength, token);

	*length = (*token & lengthMask) + 3;

	unsigned __int16 tmp1 = *token & offsetMask;
	unsigned __int16 tmp2 = 16 - bitCount;

	*offset = (tmp1 >> tmp2) + 1;
}

void decompressAToken(MS_OVBA_state *state, int i, unsigned char *b)
{
	unsigned char flag = (*b >> i) & 0x01;

	if (flag == 0)
	{
		state->dbs.decompressedCurrent[0] = state->ccs.compressedCurrent[0];
		state->dbs.decompressedCurrent++;
		state->ccs.compressedCurrent++;
	}
	else
	{
		unsigned __int16 *token = (unsigned __int16 *)state->ccs.compressedCurrent;

		unsigned __int16 offset;
		unsigned __int16 length;

		unpackCopyToken(state, &offset, &length, token);

		unsigned char *copySource = state->dbs.decompressedCurrent - offset;

		// can't use memcpy, since it can not be optimized!
		byteCopy(state->dbs.decompressedCurrent, copySource, length);

		state->dbs.decompressedCurrent += length;
		state->ccs.compressedCurrent += 2;

	}
}

void decompressTokenSequence(MS_OVBA_state *state, unsigned char *compressedEnd)
{
	unsigned char *b = state->ccs.compressedCurrent;
	state->ccs.compressedCurrent++;

	if (state->ccs.compressedCurrent < compressedEnd)
	{
		for (int i = 0; i < 8; i++)
		{
			if (state->ccs.compressedCurrent < compressedEnd)
			{
				decompressAToken(state, i, b);
			}
		}
	}
}

void decompressRawChunk(MS_OVBA_state *state)
{
	memcpy(state->dbs.decompressedCurrent, state->ccs.compressedCurrent, 4096);
	state->dbs.decompressedCurrent += 4096;
	state->ccs.compressedCurrent += 4096;
}

void decompressCompressedChunk(MS_OVBA_state *state)
{
	CompressedChunk *cc;

	cc = (CompressedChunk *)state->cchs.compressedChunkStart;

	state->dchs.decompressedChunkStart = state->dbs.decompressedCurrent;

	unsigned char *compressedEnd = MIN(state->ccs.compressedRecordEnd, state->dchs.decompressedChunkStart + cc->header.compressedChunkSize);

	state->ccs.compressedCurrent = state->cchs.compressedChunkStart + 2;

	if (cc->header.compressedChunkFlag == 0x1)
	{
		while (state->ccs.compressedCurrent < compressedEnd)
		{
			decompressTokenSequence(state, compressedEnd);
		}
	}
	else
	{
		decompressRawChunk(state);
	}
}

MS_OVBA_out MS_OVBA_decompress(void *data, size_t length) 
{
	MS_OVBA_out result = {0, nullptr, 0};
	MS_OVBA_state state;

	state.ccs.compressedCurrent = (unsigned char *)data;
	state.ccs.compressedRecordEnd = &((unsigned char *)data)[length];

	state.db = nullptr;
	state.dbs.decompressedBufferEnd = nullptr;
	state.dbs.decompressedCurrent = nullptr;

	if (state.ccs.compressedCurrent[0] == 0x01)
	{
		state.ccs.compressedCurrent += 1;

		while (state.ccs.compressedCurrent < state.ccs.compressedRecordEnd)
		{
			size_t currentSize = DECOMPRESSED_BUFFER_SIZE(state);
			state.db = (unsigned char *)realloc(state.db, currentSize + 4096);
			
			if (state.db == nullptr) 
			{
				// not enough resources
				result.state = MS_OVBA_OUT_OF_MEMORY;
				return result;
			}

			state.dbs.decompressedBufferEnd = state.db + 4096;
			state.dbs.decompressedCurrent = state.db + currentSize;

			state.cchs.compressedChunkStart = state.ccs.compressedCurrent;
			decompressCompressedChunk(&state);
		}

	}
	else
	{
		// invalid signature
		result.state = MS_OVBA_INVALID_SIGNATURE;
	}

	result.data = state.db;
	result.length = DECOMPRESSED_BUFFER_SIZE(state);

	return result;
}

void MS_OVBA_free(MS_OVBA_out r)  
{
	if (r.state != MS_OVBA_VALID && r.data != nullptr)
	{
		free(r.data);
	}

	// invalidate structure
	r.state = MS_OVBA_INVALID;
	r.data = nullptr;
	r.length = 0;
}