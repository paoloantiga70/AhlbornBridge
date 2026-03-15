#pragma once

#define LOAD_FAVORITE_ORGAN_1   1
#define LOAD_FAVORITE_ORGAN_2   2
#define LOAD_FAVORITE_ORGAN_3   3
#define LOAD_FAVORITE_ORGAN_4   4
#define LOAD_FAVORITE_ORGAN_5   5
#define LOAD_FAVORITE_ORGAN_6   6

// Sends MIDI message BF 50 00 (CC ch16, CC#80, value 0) on the current output device.
void SendUnloadOrganMidiMessage();