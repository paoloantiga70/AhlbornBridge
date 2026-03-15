#include "StreamDeck.h"
#include "Midi.h"

void SendUnloadOrganMidiMessage()
{
    // BF 50 00 = CC ch16, CC#80, value 0
    DWORD msg = CC_ch16 | (BF_0x50 << 8) | (0x00 << 16);
    if (EnqueueMidiOutMessage(msg))
    {
        printf("\nSendUnloadOrganMidiMessage: enqueued BF 50 00.\n\n");
    }
    else
    {
        printf("\nSendUnloadOrganMidiMessage: failed to enqueue message.\n\n");
    }
}
