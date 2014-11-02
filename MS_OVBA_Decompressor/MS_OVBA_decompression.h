#include <errno.h>

#define MS_OVBA_VALID 0
#define MS_OVBA_INVALID 1
#define MS_OVBA_INVALID_SIGNATURE 2
#define MS_OVBA_OUT_OF_MEMORY 3




typedef struct 
{
	size_t length;
	void *data;
	int state;
} MS_OVBA_out;

MS_OVBA_out MS_OVBA_decompress(void *, size_t);

void MS_OVBA_free(MS_OVBA_out);