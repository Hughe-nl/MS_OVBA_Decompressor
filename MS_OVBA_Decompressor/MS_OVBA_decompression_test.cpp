#include "stdafx.h"
#include "MS_OVBA_decompression.h"

/* prototype */
void MS_OVBA_decompression_test(int, unsigned char *, size_t, char *);

#define INFO 0
#define ERROR 1
#define SUCCESS 1

#define PRINT_INFO(...) if (INFO) printf( __VA_ARGS__)
#define PRINT_ERROR(...) if (ERROR) printf( __VA_ARGS__)
#define PRINT_SUCCESS(...) if (SUCCESS) printf( __VA_ARGS__)

// Test case: http://msdn.microsoft.com/en-us/library/dd944217%28v=office.12%29.aspx
void MS_OVBA_decompression_test_3()
{
	unsigned char input[] = { 0x01, 0x03, 0xB0, 0x02, 0x61, 0x45, 0x00 };
	char *expected = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";

	MS_OVBA_decompression_test(3, input, sizeof(input), expected);
}


// Test case: http://msdn.microsoft.com/en-us/library/dd946387%28v=office.12%29.aspx
void MS_OVBA_decompression_test_2()
{
	unsigned char input[] = { 0x01, 0x19, 0xB0, 0x00, 0x61, 0x62, 0x63, 0x64, 0x65, 0x66, 0x67, 0x68, 0x00, 0x69, 0x6A, 0x6B, 0x6C, 0x6D, 0x6E, 0x6F, 0x70, 0x00, 0x71, 0x72, 0x73, 0x74, 0x75, 0x76, 0x2E };
	char *expected = "abcdefghijklmnopqrstuv.";

	MS_OVBA_decompression_test(2, input, sizeof(input), expected);
}

// Test case: http://msdn.microsoft.com/en-us/library/dd773050%28v=office.12%29.aspx
void MS_OVBA_decompression_test_1()
{

	unsigned char input[] = { 0x01, 0x2F, 0xB0, 0x00, 0x23, 0x61, 0x61, 0x61, 0x62, 0x63, 0x64, 0x65, 0x82, 0x66, 0x00, 0x70, 0x61, 0x67, 0x68, 0x69, 0x6A, 0x01, 0x38, 0x08, 0x61, 0x6B, 0x6C, 0x00, 0x30, 0x6D, 0x6E, 0x6F, 0x70, 0x06, 0x71, 0x02, 0x70, 0x04, 0x10, 0x72, 0x73, 0x74, 0x75, 0x76, 0x10, 0x77, 0x78, 0x79, 0x7A, 0x00, 0x3C };
	char *expected = "#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa";

	MS_OVBA_decompression_test(1, input, sizeof(input), expected);

}

void MS_OVBA_decompression_test(int testid, unsigned char *input, size_t size, char *expected)
{

	PRINT_INFO("[*] MS_OVBA_decompression test (%d) started\r\n", testid);

	MS_OVBA_out output = MS_OVBA_decompress(input, size);

	// Check Error code
	if (output.state == MS_OVBA_VALID)
	{
		PRINT_INFO("[*] MS_OVBA_decompression valid\r\n");
	}
	else
	{
		PRINT_ERROR("[!] MS_OVBA_decompression error\r\n");
		goto onerror;
	}

	// Check Length
	if (strlen(expected) == output.length)
	{
		PRINT_INFO("[*] MS_OVBA_decompression length is correct\r\n");
	}
	else
	{
		PRINT_ERROR("[!] MS_OVBA_decompression length is incorrect\r\n");
		goto onerror;
	}

	// Check Content
	if (memcmp(output.data, expected, output.length) == 0)
	{
		PRINT_INFO("[*] MS_OVBA_decompression data is correct\r\n");
	}
	else
	{
		PRINT_ERROR("[!] MS_OVBA_decompression data is incorrect\r\n");
		goto onerror;
	}

	MS_OVBA_free(output);
	PRINT_SUCCESS("[*] MS_OVBA_decompression_test (%d) SUCCESS\r\n", testid);

	return;

onerror:
	MS_OVBA_free(output);
	PRINT_ERROR("[!] MS_OVBA_decompression_test (%d) FAILED\r\n", testid);

}
