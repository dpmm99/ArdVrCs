#include <Windows.h>
#define MESSAGE_HEADER_SIZE 2
#define MESSAGE_HEADER_BYTE 0xFE
#define MAX_INPUTS 4
#define MESSAGE_LENGTH (MAX_INPUTS + MESSAGE_HEADER_SIZE)
#define BUFFER_SIZE (MESSAGE_LENGTH * 2)
#define INPUT_THRESHOLD 95
#define INPUT_RELEASE_THRESHOLD 55
NTSTATUS(__stdcall *NtDelayExecution)(BOOL Alertable, PLARGE_INTEGER DelayInterval);
NTSTATUS(__stdcall *ZwSetTimerResolution)(IN ULONG RequestedResolution, IN BOOLEAN Set, OUT PULONG ActualResolution);

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
	WORD vkey;
	byte lastVal; //ArdVrKs (C#) uses an array so you can look at the differential. I may want to do that here, too, but getting the release timing perfect isn't as important (and my FSRs are rated for a 15ms release anyway, or 66.67 times per second per arrow...that's a big physical no).
	byte state;
	byte stateSimulated;
} INPUT_STATE;

void updateInputState(INPUT_STATE *input, byte newVal) {
	input->lastVal = newVal;

	//Determine new state
	if (input->state == 0) //Not pressed
	{
		if (input->lastVal >= INPUT_THRESHOLD) input->state = 1;
	}
	else if (input->state == 1) //Just pressed
	{
		input->state = 2;
	}
	else if (input->state == 2) //Held
	{
		if (input->lastVal <= INPUT_RELEASE_THRESHOLD) input->state = 3;
	}
	else if (input->state == 3) //Just released
	{
		input->state = 0;
	}
}

void main() {
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

	//Prepare my four input states, in the same order as the pressure data in the incoming messages
	INPUT_STATE inputs[MAX_INPUTS] = { {.vkey = 'E'}, {.vkey = 'A'}, {.vkey = 'O'}, {.vkey = VK_OEM_COMMA} }; //Dvorak equivalent of DASW

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

			//Rotate the consumed data out of the buffer
			bufferBytesFilled -= MESSAGE_LENGTH;
			memmove(buffer, buffer + MESSAGE_LENGTH, bufferBytesFilled);

			if (focusedInStepMania()) {
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

			//The delay here actually isn't necessary as long as ReadFile has a nonzero timeout.
			microsleep(500); //500 microseconds = half a millisecond = very slightly below 2k executions per second (depends on the time it takes to execute)
		}
	}

	//if (serialPort != INVALID_HANDLE_VALUE) CloseHandle(serialPort);
}