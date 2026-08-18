Attribute VB_Name = "MIDCONST"
Option Explicit
'
' State Constants
'
Global Const MIDISTATE_CLOSED = 0
Global Const MIDISTATE_OPEN = 1
Global Const MIDISTATE_STARTED = 2
Global Const MIDISTATE_STOPPED = 3
Global Const MIDISTATE_PAUSED = 4
'
' MidiOut actions
'
Global Const MIDIOUT_NONE = 0
Global Const MIDIOUT_OPEN = 1
Global Const MIDIOUT_CLOSE = 2
Global Const MIDIOUT_RESET = 3
Global Const MIDIOUT_START = 4
Global Const MIDIOUT_STOP = 5
Global Const MIDIOUT_QUEUE = 6
Global Const MIDIOUT_SEND = 7
Global Const MIDIOUT_TIMER = 8
Global Const MIDIOUT_PAUSE = 9
'
' MidiIn actions
'
Global Const MIDIIN_NONE = 0
Global Const MIDIIN_OPEN = 1
Global Const MIDIIN_CLOSE = 2
Global Const MIDIIN_RESET = 3
Global Const MIDIIN_START = 4
Global Const MIDIIN_STOP = 5
Global Const MIDIIN_REMOVE = 6
'
' MidiOut device types
'
Global Const MIDIOUT_PORT = 0
Global Const MIDIOUT_SQUARESYNTH = 1
Global Const MIDIOUT_FMSYNTH = 2
Global Const MIDIOUT_MIDIMAPPER = 3
'
' MidiFile actions
'
Global Const MIDIFILE_NONE = 0
Global Const MIDIFILE_OPEN = 1
Global Const MIDIFILE_CLOSE = 2
Global Const MIDIFILE_CREATE = 3
Global Const MIDIFILE_SAVE = 4
Global Const MIDIFILE_CLEAR = 5
Global Const MIDIFILE_INSERT_MESSAGE = 6
Global Const MIDIFILE_MODIFY_MESSAGE = 7
Global Const MIDIFILE_DELETE_MESSAGE = 8
Global Const MIDIFILE_INSERT_TRACK = 9
Global Const MIDIFILE_DELETE_TRACK = 10
Global Const MIDIFILE_SAVE_AS = 11
'
' Standard MIDI File Meta Event Constants
'
Global Const META = &HFF
Global Const META_SEQUENCE_NUMBER = &H0
Global Const META_TEXT = &H1
Global Const META_COPYRIGHT = &H2
Global Const META_NAME = &H3
Global Const META_INST_NAME = &H4
Global Const META_LYRIC = &H5
Global Const META_MARKER = &H6
Global Const META_CUE_POINT = &H7
Global Const META_CHAN_PREFIX = &H20
Global Const META_EOT = &H2F
Global Const META_TEMPO = &H51
Global Const META_SMPTE_OFFSET = &H54
Global Const META_TIME_SIG = &H58
Global Const META_KEY_SIG = &H59
Global Const META_SEQ_SPECIFIC = &H7F
'
' MIDI System Command Filter
'
Global Const FILTER_MTC = &HF1          'filter MTC
Global Const FILTER_SNGPTR = &HF2       'filter Song Position Pointer F2
Global Const FILTER_SNGSEL = &HF3       'filter Song Select F3
Global Const FILTER_F4 = &HF4           'filter undefined F4
Global Const FILTER_F5 = &HF5           'filter undefined F5
Global Const FILTER_TUNE = &HF6         'filter Tune Request F6
Global Const FILTER_CLOCK = &HF8        'filter MIDI Clock F8
Global Const FILTER_F9 = &HF9           'filter undefined F9
Global Const FILTER_START = &HFA        'filter MIDI Start FA
Global Const FILTER_CONT = &HFB         'filter MIDI Continue FB
Global Const FILTER_STOP = &HFC         'filter MIDI Stop FC
Global Const FILTER_FD = &HFD           'filter undefined FD
Global Const FILTER_ACTIVE_SENSE = &HFE 'filter Active Sensing FE
Global Const FILTER_RESET = &HFF        'filter System Reset FF
'
' MIDI status messages
'
Global Const NOTE_OFF = &H80
Global Const NOTE_ON = &H90
Global Const POLY_KEY_PRESS = &HA0
Global Const CONTROLLER_CHANGE = &HB0
Global Const PROGRAM_CHANGE = &HC0
Global Const CHANNEL_PRESSURE = &HD0
Global Const PITCH_BEND = &HE0
Global Const SYSEX = &HF0
Global Const MTC_QFRAME = &HF1
Global Const EOX = &HF7
Global Const MIDI_CLOCK = &HF8
Global Const MIDI_START = &HFA
Global Const MIDI_CONTINUE = &HFB
Global Const MIDI_STOP = &HFC
'
' CONTROLLER_CHANGE, MIDI Controller Numbers Constants
'
Global Const MOD_WHEEL = 1
Global Const BREATH_CONTROLLER = 2
Global Const FOOT_CONTROLLER = 4
Global Const PORTAMENTO_TIME = 5
Global Const MAIN_VOLUME = 7
Global Const BALANCE = 8
Global Const PAN = 10
Global Const EXPRESS_CONTROLLER = 11
Global Const DAMPER_PEDAL = 64
Global Const PORTAMENTO = 65
Global Const SOSTENUTO = 66
Global Const SOFT_PEDAL = 67
Global Const HOLD_2 = 69
Global Const EXTERNAL_FX_DEPTH = 91
Global Const TREMELO_DEPTH = 92
Global Const CHORUS_DEPTH = 93
Global Const DETUNE_DEPTH = 94
Global Const PHASER_DEPTH = 95
Global Const DATA_INCREMENT = 96
Global Const DATA_DECREMENT = 97
'
' Constants to simplify MIDI
'
Global Const TOTAL_MIDI_CHANNELS = 16
'
' MIDI hangkészlet
'
Global Hang(128) As String

Public Sub MIDI_Hangok()
 '
 ' MIDI hangkészlet
 '
 Hang(1) = "Acoustic Grand Piano"
 Hang(2) = "Bright Acoustic Piano"
 Hang(3) = "Electric Grand Piano"
 Hang(4) = "Honkytonk Piano"
 Hang(5) = "Rhodes Piano"
 Hang(6) = "Chorus Piano"
 Hang(7) = "Harpsichord"
 Hang(8) = "Clavinet"
 Hang(9) = "Celesta"
 Hang(10) = "Glockenspiel"
 Hang(11) = "Music Box"
 Hang(12) = "Vibraphone"
 Hang(13) = "Marimba"
 Hang(14) = "Xylophone"
 Hang(15) = "Tubular Bells"
 Hang(16) = "Dulcimer"
 Hang(17) = "Hammond Organ"
 Hang(18) = "Percuss. Organ"
 Hang(19) = "Rock Organ"
 Hang(20) = "Church Organ"
 Hang(21) = "Reed Organ"
 Hang(22) = "Accordion"
 Hang(23) = "Harmonica"
 Hang(24) = "Tango Accordion"
 Hang(25) = "Acoustic Guitar (nylon)"
 Hang(26) = "Acoustic Guitar (steel)"
 Hang(27) = "Electric Guitar (jazz)"
 Hang(28) = "Electric Guitar (clean)"
 Hang(29) = "Electric Guitar (muted)"
 Hang(30) = "Overdriven Guitar"
 Hang(31) = "Distortion Guitar"
 Hang(32) = "Guitar Harmonics"
 Hang(33) = "Acoustic Bass"
 Hang(34) = "Electric Bass (finger)"
 Hang(35) = "Electric Bass (pick)"
 Hang(36) = "Fretless Bass"
 Hang(37) = "Slap Bass 1"
 Hang(38) = "Slap Bass 2"
 Hang(39) = "Synth Bass 1"
 Hang(40) = "Synth Bass 2"
 Hang(41) = "Violin"
 Hang(42) = "Viola"
 Hang(43) = "Cello"
 Hang(44) = "Contra Bass"
 Hang(45) = "Tremolo Strings"
 Hang(46) = "Pizzicato Strings"
 Hang(47) = "Orchestral Harp"
 Hang(48) = "Timpani"
 Hang(49) = "String Ensemble 1"
 Hang(50) = "String Ensemble 2"
 Hang(51) = "Synth Strings 1"
 Hang(52) = "Synth Strings 2"
 Hang(53) = "Choir Aahs"
 Hang(54) = "Voice Oohs"
 Hang(55) = "Synth Voice"
 Hang(56) = "Orchestra Hit"
 Hang(57) = "Trumpet"
 Hang(58) = "Trombone"
 Hang(59) = "Tuba"
 Hang(60) = "Muted Trumpet"
 Hang(61) = "French Horn"
 Hang(62) = "Brass Section"
 Hang(63) = "Synth Brass 1"
 Hang(64) = "Synth Brass 2"
 Hang(65) = "Soprano Sax"
 Hang(66) = "Alto Sax"
 Hang(67) = "Tenor Sax"
 Hang(68) = "Baritone Sax"
 Hang(69) = "Oboe"
 Hang(70) = "English Horn"
 Hang(71) = "Bassoon"
 Hang(72) = "Clarinet"
 Hang(73) = "Piccolo"
 Hang(74) = "Flute"
 Hang(75) = "Recorder"
 Hang(76) = "Pan Flute"
 Hang(77) = "Bottle Blow"
 Hang(78) = "Shaku"
 Hang(79) = "Whistle"
 Hang(80) = "Ocarina"
 Hang(81) = "Lead 1 (square)"
 Hang(82) = "Lead 2 (saw tooth)"
 Hang(83) = "Lead 3 (calliope lead)"
 Hang(84) = "Lead 4 (chiff lead)"
 Hang(85) = "Lead 5 (charang)"
 Hang(86) = "Lead 6 (voice)"
 Hang(87) = "Lead 7 (fifths)"
 Hang(88) = "Lead 8 (bass + lead)"
 Hang(89) = "Pad 1 (new age)"
 Hang(90) = "Pad 2 (warm)"
 Hang(91) = "Pad 3 (poly synth)"
 Hang(92) = "Pad 4 (choir)"
 Hang(93) = "Pad 5 (bowed)"
 Hang(94) = "Pad 6 (metallic)"
 Hang(95) = "Pad 7 (halo)"
 Hang(96) = "Pad 8 (sweep)"
 Hang(97) = "FX 1 (rain)"
 Hang(98) = "FX 2 (sound track)"
 Hang(99) = "FX 3 (crystal)"
 Hang(100) = "FX 4 (atmosphere)"
 Hang(101) = "FX 5 (bright)"
 Hang(102) = "FX 6 (goblins)"
 Hang(103) = "FX 7 (echoes)"
 Hang(104) = "FX 8 (sci-fi)"
 Hang(105) = "Sitar"
 Hang(106) = "Banjo"
 Hang(107) = "Shamisen"
 Hang(108) = "Koto"
 Hang(109) = "Kalimba"
 Hang(110) = "Bagpipe"
 Hang(111) = "Fiddle"
 Hang(112) = "Shanai"
 Hang(113) = "Tinkle Bell"
 Hang(114) = "Agogo"
 Hang(115) = "Steel Drums"
 Hang(116) = "Wood block"
 Hang(117) = "Taiko Drum"
 Hang(118) = "Melodic Tom"
 Hang(119) = "Synth Drum"
 Hang(120) = "Reverse Cymbal"
 Hang(121) = "Guitar Fret Noise"
 Hang(122) = "Breath Noise"
 Hang(123) = "Seashore"
 Hang(124) = "Bird Tweet"
 Hang(125) = "Telephone Ring"
 Hang(126) = "Helicopter"
 Hang(127) = "Applause"
 Hang(128) = "Gunshot"

End Sub
