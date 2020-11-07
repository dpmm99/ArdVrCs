#include <Windows.h>
#define MESSAGE_HEADER_SIZE 2
#define MESSAGE_HEADER_BYTE 0xFE
#define MAX_INPUTS 4
#define MESSAGE_LENGTH (MAX_INPUTS + MESSAGE_HEADER_SIZE)
#define BUFFER_SIZE (MESSAGE_LENGTH * 2)
#define BYTES_PER_INPUT_CONFIG 3
#define CONFIG_FILE_SIZE (MAX_INPUTS * BYTES_PER_INPUT_CONFIG)
//#define TEST_FILE_MODE
#include <stdio.h>
#include <time.h>
#include <conio.h>

NTSTATUS(__stdcall *NtDelayExecution)(BOOL Alertable, PLARGE_INTEGER DelayInterval);
NTSTATUS(__stdcall *ZwSetTimerResolution)(IN ULONG RequestedResolution, IN BOOLEAN Set, OUT PULONG ActualResolution);

#ifdef TEST_FILE_MODE
FILE *outFile = NULL;
int fileWrites = 0;
void writeToFile(const void *bytes) { //Pass in a byte for each input.
	if (!outFile && fileWrites < 10000) { //No more than ~10 seconds of data
		fopen_s(&outFile, "test.bin", "wb");
	}
	else if (outFile && fileWrites == 10000) {
		fclose(outFile);
	}

	if (outFile) {
		fwrite(bytes, MAX_INPUTS, 1, outFile);
	}
}
#endif

int focusedInStepMania() {
	HWND foregroundWindow = GetForegroundWindow();
	char strBuffer[256];
	int titleCharCount = GetWindowText(foregroundWindow, strBuffer, sizeof(strBuffer));
	if (strstr(strBuffer, "Simply Love") != NULL) { //If our foreground window has this anywhere in the title
		return -1;
	}
	return 0;
}

void microsleep(int microseconds) {
	LARGE_INTEGER interval;
	interval.QuadPart = -1 * (int)(microseconds * 10);
	NtDelayExecution(FALSE, &interval);
}

void sendKey(WORD vkey, DWORD flags) { //flags is... IsExtendedKey(keyCode) ? KEYEVENTF_EXTENDEDKEY : 0... and up, down, left, right are extended keys. For key release: KEYEVENTF_KEYUP
	INPUT inputs[] = {
		{
			.type = INPUT_KEYBOARD,
			.ki = {
				.wVk = vkey,
				.dwFlags = flags
			}
		}
	};
	inputs[0].ki.wScan = MapVirtualKey(inputs[0].ki.wVk, 0);
	SendInput(1, inputs, sizeof(INPUT));
}

HANDLE openSerial() {
	HANDLE comHandle;
	comHandle = CreateFile("COM4", GENERIC_READ, 0 /*no sharing*/, NULL /*no security*/, OPEN_EXISTING, 0 /*non-overlapped I/O*/, NULL);

	if (comHandle == INVALID_HANDLE_VALUE) return comHandle;

	DCB comSettings = {
		.DCBlength = sizeof(DCB)
	};
	if (GetCommState(comHandle, &comSettings) == FALSE) { //No point in keeping the port open if we can't get/change the settings
		CloseHandle(comHandle);
		return INVALID_HANDLE_VALUE;
	}

	comSettings.BaudRate = CBR_115200; //10 bits per byte (1 start bit, 1 stop bit), 6 bytes per message, 60 bits per message, 1920 messages per second this way
	comSettings.ByteSize = 8;
	comSettings.StopBits = ONESTOPBIT;
	comSettings.Parity = NOPARITY;
	if (SetCommState(comHandle, &comSettings) == FALSE) {
		CloseHandle(comHandle);
		return INVALID_HANDLE_VALUE;
	}

	//Set all the timeouts to 1ms because a complete message, as stated above, should take less than 1ms to arrive.
	COMMTIMEOUTS timeoutSettings = {
		.ReadIntervalTimeout = 1, //Max wait time from one byte to the next
		.ReadTotalTimeoutMultiplier = 1,
		.ReadTotalTimeoutConstant = 1,
		.WriteTotalTimeoutMultiplier = 1,
		.WriteTotalTimeoutConstant = 1
	};
	if (SetCommTimeouts(comHandle, &timeoutSettings) == FALSE) { //It's totally feasible for that port to have died in the last couple nanoseconds. :P
		CloseHandle(comHandle);
		return INVALID_HANDLE_VALUE;
	}

	return comHandle;
}

typedef struct INPUT_STATE {
	byte vkey; //vkey is a WORD (unsigned short) but the documentation says it has to be between 0 and 254 inclusive
	byte lastVal; //ArdVrKs (C#) uses an array so you can look at the differential. I may want to do that here, too, but getting the release timing perfect isn't as important (and my FSRs are rated for a 15ms release anyway, or 66.67 times per second per arrow...that's a big physical no).
	byte state;
	byte stateSimulated;

	byte pressThreshold;
	byte releaseThreshold;
} INPUT_STATE;

void updateInputState(INPUT_STATE *input, byte newVal) {
	input->lastVal = newVal;

	//Determine new state
	if (input->state == 0) //Not pressed
	{
		if (input->lastVal >= input->pressThreshold) input->state = 1;
	}
	else if (input->state == 1) //Just pressed
	{
		input->state = 2;
	}
	else if (input->state == 2) //Held
	{
		if (input->lastVal <= input->releaseThreshold) input->state = 3;
	}
	else if (input->state == 3) //Just released
	{
		input->state = 0;
	}
}

void loadConfig(INPUT_STATE *inputs) {
	//Open config file
	FILE *inFile;
	errno_t openError;
	if (openError = fopen_s(&inFile, "config.bin", "rb")) {
		switch (openError) {
		case ENOENT: //https://docs.microsoft.com/en-us/cpp/c-runtime-library/errno-doserrno-sys-errlist-and-sys-nerr?view=msvc-160
			printf("config.bin not found. Using default keys and thresholds.\n");
			break;
		default:
			printf("Unknown error %d when opening config.bin. Using default configuration.\n", openError);
		}
		return;
	}

	//Read config file
	byte buffer[CONFIG_FILE_SIZE];
	size_t bytes_read;
	bytes_read = fread(buffer, sizeof(byte), CONFIG_FILE_SIZE, inFile);
	if (bytes_read < CONFIG_FILE_SIZE) {
		printf("Unexpected end of file. Using default thresholds.\n");
		fclose(inFile);
		return;
	}

	fclose(inFile);

	//Apply config
	for (int x = 0; x < MAX_INPUTS; x++) {
		inputs[x].vkey = buffer[x * BYTES_PER_INPUT_CONFIG];
		inputs[x].pressThreshold = buffer[x * BYTES_PER_INPUT_CONFIG + 1];
		inputs[x].releaseThreshold = buffer[x * BYTES_PER_INPUT_CONFIG + 2];
	}

	printf("Configuration loaded.\n");
}

void saveConfig(INPUT_STATE *inputs) {
	//Make backup if the file exists. (Fails silently, but just about the only error case in which it may lead to the old config.bin being lost is when the file path is too long.)
	char backupFilename[32];
	time_t now = time(NULL);
	struct tm nowLocal;
	localtime_s(&nowLocal, &now);
	strftime(backupFilename, sizeof(backupFilename), "config_%Y-%m-%d_%H-%M-%S.bin", &nowLocal);
	rename("config.bin", backupFilename);

	//Create config file
	FILE *outFile;
	errno_t openError;
	if (openError = fopen_s(&outFile, "config.bin", "wb")) {
		switch (openError) {
		case EROFS: //https://docs.microsoft.com/en-us/cpp/c-runtime-library/errno-doserrno-sys-errlist-and-sys-nerr?view=msvc-160
			printf("Filesystem appears to be read-only. Cannot save config.bin.\n");
			break;
		default:
			printf("Unknown error %d when creating file. Cannot save config.bin.\n", openError);
		}
		return;
	}

	//Save config
	for (int x = 0; x < MAX_INPUTS; x++) {
		fwrite(&inputs[x].vkey, sizeof(byte), 1, outFile);
		fwrite(&inputs[x].pressThreshold, sizeof(byte), 1, outFile);
		fwrite(&inputs[x].releaseThreshold, sizeof(byte), 1, outFile);
	}

	fclose(outFile);

	printf("Configuration saved.\n");
}

char *getNameFromVirtualKey(byte vkey) {
	static char keyName[32] = { 0 }; //Not thread safe. :)
	WORD scanCode = MapVirtualKey(vkey, 0);

	//"because MapVirtualKey strips the extended bit for some keys", according to http://www.setnode.com/blog/mapvirtualkey-getkeynametext-and-a-story-of-how-to/
	switch (vkey)
	{
	case VK_LEFT: case VK_UP: case VK_RIGHT: case VK_DOWN:
	case VK_PRIOR: case VK_NEXT: //Page up/down
	case VK_END: case VK_HOME:
	case VK_INSERT: case VK_DELETE:
	case VK_DIVIDE: //Numpad slash
	case VK_NUMLOCK:
		scanCode |= 0x100; //Set 'extended' flag bit
	}

	GetKeyNameText(scanCode << 16, keyName, sizeof(keyName));

	return keyName;
}

void continueCalibrating(int *calibrating, INPUT_STATE *inputs) {
	static byte minima[MAX_INPUTS] = { 0 };
	static byte maxima[MAX_INPUTS] = { 0 };
	static byte displayCharIndexes[MAX_INPUTS] = { 0 };
	int x;

	//Calibrate and update inputs. then calculate the final numbers, then let them test it out, and if they press escape, restart calibration, and if they press any other key, save the file and end calibration mode.
	(*calibrating)++; //Updated up to about 2000 times per second in theory, but the Arduino program is set to closer to half of that
	if ((*calibrating >= 100 && *calibrating < 1000) || (*calibrating >= 4000 && *calibrating < 5000) || (*calibrating >= 8000 && *calibrating < 9000)) { //Dead zones from 1 to 100, 1000 to 4000, etc. to avoid reading garbage inputs and to give the user time to move
		if (*calibrating == 100 || *calibrating == 4000 || *calibrating == 8000) { //Reset the minima to the max value and the maxima to the min value before each recording
			memset(minima, 255, sizeof(minima));
			memset(maxima, 0, sizeof(maxima));
		}

		for (x = 0; x < MAX_INPUTS; x++) {
			if (minima[x] > inputs[x].lastVal) minima[x] = inputs[x].lastVal;
			if (maxima[x] < inputs[x].lastVal) maxima[x] = inputs[x].lastVal;
		}
	}
	else if (*calibrating == 1000) {
		//Save the maxima as tentative release thresholds
		for (x = 0; x < MAX_INPUTS; x++) inputs[x].releaseThreshold = maxima[x];

		printf("Please quickly stand on the left and right arrows and wait a few seconds.\n");
	}
	else if (*calibrating == 5000) { //These two blocks are going to be totally wrong if you set up the inputs in a different order. :P
		inputs[0].pressThreshold = minima[0]; //Right
		inputs[1].pressThreshold = minima[1]; //Left
		if (maxima[2] > inputs[2].releaseThreshold) inputs[2].releaseThreshold = maxima[2]; //Down
		if (maxima[3] > inputs[3].releaseThreshold) inputs[3].releaseThreshold = maxima[3]; //Up

		printf("Please quickly stand on the up and down arrows and wait a few seconds.\n");
	}
	else if (*calibrating == 9000) {
		if (maxima[0] > inputs[0].releaseThreshold) inputs[0].releaseThreshold = maxima[0]; //Right
		if (maxima[1] > inputs[1].releaseThreshold) inputs[1].releaseThreshold = maxima[1]; //Left
		inputs[2].pressThreshold = minima[2]; //Down
		inputs[3].pressThreshold = minima[3]; //Up

		//Now we have measured actual press/release thresholds, but we can drop the press threshold down if it's much higher than the release threshold, and it might even be okay for us to push the release threshold down!
		for (x = 0; x < MAX_INPUTS; x++) {
			if ((int)inputs[x].pressThreshold > (int)inputs[x].releaseThreshold + 5) inputs[x].pressThreshold = inputs[x].releaseThreshold + 1;
			else if (inputs[x].releaseThreshold > 5) { //They're already pretty close together; let's try to gain some leeway by moving the release threshold down a bit, even if it causes the release to be delayed a few milliseconds due to static electricity--I find it unlikely that two notes on the same arrow will ever need to be played under 30ms apart (64th note at 300 BPM)
				inputs[x].releaseThreshold -= 5;
				inputs[x].pressThreshold -= 5;
			}
		}

		printf("Calibration results: ");
		for (x = 0; x < MAX_INPUTS; x++) {
			printf("key %s: %d/%d%s", getNameFromVirtualKey(inputs[x].vkey), inputs[x].pressThreshold, inputs[x].releaseThreshold, x < MAX_INPUTS - 1 ? ", " : "");
		}

		//Set character indexes for displaying a symbol when an arrow is stepped on
		displayCharIndexes[0] = 14; //Right
		displayCharIndexes[1] = 2; //Left
		displayCharIndexes[2] = 6; //Down
		displayCharIndexes[3] = 10; //Up
		printf("Step on the arrows to test the calibration results, then step off and press escape to restart calibration or any other key to save and continue.\nL   D   U   R   "); //No ending linefeed, so we can use absolute horizontal positioning commands to quickly update the display
	}
	else if (*calibrating == 9001) {
		//Display what inputs are currently being pressed
		for (x = 0; x < MAX_INPUTS; x++) {
			if ((inputs[x].state == 1 || inputs[x].state == 2) && inputs[x].stateSimulated != 1) {
				printf("\x1b[%dGX", displayCharIndexes[x]); //Horizontal positioning character code sequence: ESC [ <n> G
				inputs[x].stateSimulated = 1;
			}
			else if ((inputs[x].state == 0 || inputs[x].state == 3) && inputs[x].stateSimulated == 1) {
				printf("\x1b[%dG ", displayCharIndexes[x]); //Clear the X
				inputs[x].stateSimulated = 0;
			}
		}

		//Wait for the user to either hold escape to restart calibration or any other key to save the config and end calibration mode
		if (_kbhit()) {
			if (_getch() == 27) { //Escape key
				*calibrating = 1;
				printf("\nRestarting calibration...\n");
			}
			else {
				printf("\nSaving calibration data...\n");
				saveConfig(inputs);
				*calibrating = 0;
			}
		}
		else (*calibrating)--;
	}
}

void main(int argc, char *argv[]) {
	//Get function addresses from ntdll so we can use timers with extreme accuracy instead of 'sleep', which likes to jump either 0ms or 17ms every time in my experience (although I've read that the default Windows quantum is 15ms).
	//Don't forget to set this thread to real-time priority in order to keep the timing as close to perfect as possible. This may actually end up being more accurate than StepMania itself, for all I know...
	NtDelayExecution = (NTSTATUS(__stdcall*)(BOOL, PLARGE_INTEGER)) GetProcAddress(GetModuleHandle("ntdll.dll"), "NtDelayExecution");
	ZwSetTimerResolution = (NTSTATUS(__stdcall*)(ULONG, BOOLEAN, PULONG)) GetProcAddress(GetModuleHandle("ntdll.dll"), "ZwSetTimerResolution");
	ULONG actualResolution;
	ZwSetTimerResolution(500, TRUE, &actualResolution); //First parameter's units is 100 nanoseconds (e.g. 10 = 1000 nanoseconds = 1 microsecond). 10000 = one millisecond.
	HANDLE serialPort = INVALID_HANDLE_VALUE;
	byte buffer[BUFFER_SIZE] = { 0 };
	int bufferBytesFilled = 0;
	int recentlyReadBytes = 0;
	int calibrating = 0;

	SetPriorityClass(GetCurrentProcess(), HIGH_PRIORITY_CLASS); //Accurate timing is more important than other processes if you're playing a game anyway. :)
	SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_HIGHEST);

	//Prepare my four input states, in the same order as the pressure data in the incoming messages
	INPUT_STATE inputs[MAX_INPUTS] = {
		{.vkey = 'E', .pressThreshold = 0x77, .releaseThreshold = 0x70}, //Right
		{.vkey = 'A', .pressThreshold = 0x6D, .releaseThreshold = 0x66}, //Left
		{.vkey = 'O', .pressThreshold = 0x60, .releaseThreshold = 0x55}, //Down
		{.vkey = VK_OEM_COMMA, .pressThreshold = 0x3D, .releaseThreshold = 0x37} //Up
	}; //Dvorak equivalent of DASW
	loadConfig(inputs);

	if (argc > 1 && argv[1][0] && (!_strcmpi(argv[1], "calibrate") || !_strcmpi(&argv[1][1], "calibrate"))) //Allow "calibrate" or "-calibrate" or "/calibrate" or even "ncalibrate" so it doesn't matter what the user is used to. :P
	{
		calibrating = 1;
	}

	//TODO: Instantaneous thresholds aren't going to work well; I need to use the delta, but even then, all of them drop at once when you press two arrows due to my added pressure... Maybe I can drop the press threshold when you start pressing an arrow. I should definitely also get higher-resistance resistors (probably about 10k ohms) to separate the inputs from each other and from ground.

	while (1) { //TODO: Maybe don't loop forever. I might make a C# UI that starts and stops this daemon gracefully...
		while (serialPort == INVALID_HANDLE_VALUE) {
			serialPort = openSerial();
			if (serialPort == INVALID_HANDLE_VALUE) microsleep(100000); //Could just use Sleep/SleepEx here, but I've got my own, so why not! 100ms sleep between port opening attempts.
		}

		while (serialPort != INVALID_HANDLE_VALUE) {
			//Read from the serial port
			if (ReadFile(serialPort, buffer + bufferBytesFilled, sizeof(buffer) - bufferBytesFilled, &recentlyReadBytes, NULL) == FALSE) {
				//Read failed. The port may have been closed or something. (I can check GetLastError() but don't really need to.)
				CloseHandle(serialPort);
				serialPort = INVALID_HANDLE_VALUE;
				bufferBytesFilled = 0; //Treat the buffer as having been emptied
				break;
			}
			bufferBytesFilled += recentlyReadBytes; //You can check recentlyReadBytes to see if anything was actually read, but it's safe to assume that something was read if ReadFile has a timeout > 0.5ms.
			//TODO: I probably want to shrink the buffer to just 1x the whole message size to reduce the delay, but I may also want to keep at least one extra byte in case the daemon falls behind...I'd probably like to set the timeouts to 0 in that case.
			//For ReadIntervalTimeout: "A value of MAXDWORD, combined with zero values for both the ReadTotalTimeoutConstant and ReadTotalTimeoutMultiplier members, specifies that the read operation is to return immediately with the bytes that have already been received, even if no bytes have been received."

			//Identify the message start bytes--it SHOULD be true that only the first time, when you're syncing initially, they won't be at the start of the buffer. I could just reserve 0xFF in the Arduino code and only have the buttons output 0x00-0xFE, though... or maybe rely on the COM port settings, to put a start and stop bit every 4 bytes instead of every 1.
			for (int x = 0; x < bufferBytesFilled - 1; x++)
			{
				if (buffer[x] == MESSAGE_HEADER_BYTE && buffer[x + 1] == MESSAGE_HEADER_BYTE)
				{
					if (x != 0) {
						bufferBytesFilled -= x;
						memmove(buffer, buffer + x, bufferBytesFilled); //memmove, not memcpy, for overlapping memory. Move the remaining filled bytes of the buffer back toward the start, to align the message header with the start of the buffer.
					}
					break;
				}
			}

			//If we have enough data for a complete message (the rest of the code could techincally be inside the above loop instead, right before the 'break')
			if (bufferBytesFilled < MESSAGE_LENGTH) continue;

			//If we have too much data, we might actually just want to skip some of it entirely. (Note: this isn't a *good* method of playing catch-up, as it limits the max catch-up rate to 2x the usual rate, but it's better than just NOT catching up.)
			if (bufferBytesFilled == BUFFER_SIZE) {
				bufferBytesFilled -= MESSAGE_LENGTH;
				memmove(buffer, buffer + MESSAGE_LENGTH, bufferBytesFilled);
			}

			//Process button data
			for (int x = 0; x < MAX_INPUTS; x++)
			{
				updateInputState(&inputs[x], buffer[x + MESSAGE_HEADER_SIZE]);
			}

#ifdef TEST_FILE_MODE
			writeToFile(buffer + MESSAGE_HEADER_SIZE); //Write those same bytes to a file
#endif

			//Rotate the consumed data out of the buffer
			bufferBytesFilled -= MESSAGE_LENGTH;
			memmove(buffer, buffer + MESSAGE_LENGTH, bufferBytesFilled);

#ifndef TEST_FILE_MODE
			if (calibrating) continueCalibrating(&calibrating, inputs);
			else if (focusedInStepMania()) {
				//Send simulated input events where the button simulated state doesn't (essentially) match the current state
				for (int x = 0; x < MAX_INPUTS; x++)
				{
					if ((inputs[x].state == 1 || inputs[x].state == 2) && inputs[x].stateSimulated != 1) {
						sendKey(inputs[x].vkey, 0);
						inputs[x].stateSimulated = 1;
					}
					else if ((inputs[x].state == 0 || inputs[x].state == 3) && inputs[x].stateSimulated == 1) {
						sendKey(inputs[x].vkey, KEYEVENTF_KEYUP);
						inputs[x].stateSimulated = 0;
					}
				}
			}
#endif
			//The delay here actually isn't necessary as long as ReadFile has a nonzero timeout.
			microsleep(500); //500 microseconds = half a millisecond = very slightly below 2k executions per second (depends on the time it takes to execute)
		}
	}

	//if (serialPort != INVALID_HANDLE_VALUE) CloseHandle(serialPort);
}