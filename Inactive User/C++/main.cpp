#define WINDOWS_LEAN_AND_MEAN
#include <windows.h>
#include <iostream>

// do something after 10 minutes of user inactivity
static const unsigned int idle_milliseconds = 60*10*1000;
// wait at least an hour between two runs
static const unsigned int interval = 60*60*1000;

int main() {
    LASTINPUTINFO last_input;
    BOOL screensaver_active;

    // main loop to check if user has been idle long enough
    for (;;) {
        if ( !GetLastInputInfo(&last_input)
          || !SystemParametersInfo(SPI_GETSCREENSAVERACTIVE, 0,  
                                   &screensaver_active, 0))
        {
            std::cerr << "WinAPI failed!" << std::endl;
            return ERROR_FAILURE;
        }

        if (last_input.dwTime < idle_milliseconds && !screensaver_active) {
            // user hasn't been idle for long enough
            // AND no screensaver is running
            Sleep(1000);
            continue;
        }

        // user has been idle at least 10 minutes
        do_something();
        // done. Wait before doing the next loop.
        Sleep(interval);
    }
}