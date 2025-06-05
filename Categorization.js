// --- CONFIGURATION ---
const OPENAI_MODEL = "gpt-4o-mini";
const OPENAI_TEMPERATURE = 0.0; 
const MAX_COMPLETION_TOKENS = 60; 
const MAX_CATEGORIES_TO_SEND = 75; 
const SHEET_NAME_PRODUCTS = "BaseTemplate";
const SHEET_NAME_CATEGORIES = "Shopify_Categories";
const TITLE_COLUMN = "A";
const DESCRIPTION_COLUMN = "T"; 
const VENDOR_COLUMN = "U";    // <<<< ADAPT THIS if your Vendor column is different
const OUTPUT_CATEGORY_COLUMN = "FY";
const MAX_API_RETRIES = 3;
const BATCH_SIZE = 50; // Maximum rows to process per batch

// --- USER TO POPULATE BASED ON CHATGPT ANALYSIS & THEIR DATA ---
const VENDOR_SPECIFIC_CATEGORY_PREFIXES = {
  "Pioneer DJ": [
    "DJ Equipment > DJ Controllers", "DJ Equipment > Mixers", "DJ Equipment > DJ Accessories",
    "DJ Equipment > DJ Cases > DJ Controller Cases", "Pro Audio > Loudspeakers", "Pro Audio > Headphones"
  ],
  "Mackie": [
    "Pro Audio > Loudspeakers", "Pro Audio > Mixers", "Pro Audio > Audio Consoles & Mixers > Audio Mixer Accessories",
    "Pro Audio > Live Sound & Portable Systems > PA Systems", "Pro Audio > Live Sound & Portable Systems > Power Amplifiers",
    "Pro Audio > Recording", "Pro Audio > Headphones"
  ],
  "Shure": ["Microphones > Wired Microphones", "Microphones > Wireless Microphones", "Microphones > Microphone Accessories", "Pro Audio > Headphones"],
  // Add many more vendor mappings here based on your ChatGPT analysis and data...
  "ExampleVendor": ["Specific Top Level > Specific SubCategory"]
};

const KEYWORD_TO_CATEGORY_GROUP_MAPPINGS = [
  {
    keywordsRegex: /speaker|loudspeaker|subwoofer|monitor\b|soundbar/i,
    targetGroupPrefixes: [
      "Pro Audio > Loudspeakers", "Home & Business AV > Home AV", "Business AV > Commercial Speakers",
      "Pro Audio > Live Sound & Portable Systems > PA Systems", "Pro Audio > Live Sound & Portable Systems > Wireless Speakers"
    ]
  },
  { keywordsRegex: /headphone|earbud|iem\b|earphones/i, targetGroupPrefixes: ["Pro Audio > Headphones", "Microphones > Wireless Microphones > Wireless IEM Systems"] },
  { keywordsRegex: /microphone|lavalier|shotgun|mic\b/i, targetGroupPrefixes: ["Microphones", "Pro Audio > Accessories > Microphone Stands", "Business AV > Audio Conferencing Systems > Conference Microphones"] },
  // ... Add more general keyword mappings ...
];

const TOP_LEVEL_CATEGORIES_FOR_FALLBACK = [
    "Business AV", "Computers & Peripherals", "DJ Equipment", "Home & Business AV",
    "Microphones", "Mobile", "Musical Instruments", "Pro Audio", "Pro Lighting", "Pro Video", "Used"
];
const DEFAULT_CATEGORY_LIST = [
  "Business AV > Assisted Listening Devices > FM Assistive Listening Devices",
  "Business AV > Assisted Listening Devices > Infrared Assistive Living Devices",
  "Business AV > Audio Conferencing Systems > Audio Controllers",
  "Business AV > Audio Conferencing Systems > Conference Microphones > Conference Room Ceiling Microphones",
  "Business AV > Audio Conferencing Systems > Conference Microphones > Conference Room Omnidirectional Microphones",
  "Business AV > Audio Conferencing Systems > Conference Microphones > Goose Conference Room Microphones",
  "Business AV > Audio Conferencing Systems > Conference Microphones > USB Conference Room Microphones",
  "Business AV > Audio Conferencing Systems > Conference Microphones > Wireless Boundary Microphones",
  "Business AV > Audio Conferencing Systems > Conference Speakerphones",
  "Business AV > Audio Conferencing Systems > Digital Lecterns",
  "Business AV > Business AV Accessories",
  "Business AV > Business Headsets",
  "Business AV > Business Intercom Systems",
  "Business AV > Commercial Audio Devices",
  "Business AV > Commercial Speakers > Commercial Ceiling Speakers",
  "Business AV > Commercial Speakers > Commercial Outdoor Speakers",
  "Business AV > Commercial Speakers > Commercial Pendant Speakers",
  "Business AV > Security and Surveillance Equipment > Keypad Readers",
  "Business AV > Security and Surveillance Equipment > Security Alarms > Security Fire Alarms",
  "Business AV > Security and Surveillance Equipment > Security Cameras",
  "Business AV > Security and Surveillance Equipment > Security Window Sensors",
  "Business AV > Security and Surveillance Equipment > Surveillance Kits",
  "Computers & Peripherals > Computer Adapters",
  "Computers & Peripherals > Computer Cables",
  "Computers & Peripherals > Computer Cables & Adapters",
  "Computers & Peripherals > Computer Cables > Computer Network Cables",
  "Computers & Peripherals > Computer Cables > Computer Network Cables > Ethernet Cables",
  "Computers & Peripherals > Computer Cables > Computer Network Cables > Optic Fiber Cables",
  "Computers & Peripherals > Computer Cables > Computer Power Cables",
  "Computers & Peripherals > Computer Cables > Computer USB Cables",
  "Computers & Peripherals > Computer Keyboards",
  "Computers & Peripherals > Computer Keyboards > Computer Keyboard Covers",
  "Computers & Peripherals > Computer Keyboards > Wired Computer Keyboards",
  "Computers & Peripherals > Computer Keyboards > Wireless Computer Keyboards",
  "Computers & Peripherals > Computer Peripherals",
  "Computers & Peripherals > Computer Peripherals > Color Calibration Tools",
  "Computers & Peripherals > Computer Peripherals > Computer Mouses",
  "Computers & Peripherals > Computer Peripherals > Gaming Racing Seats",
  "Computers & Peripherals > Computer Peripherals > Tablet Pens",
  "Computers & Peripherals > Computer Peripherals > VR Headsets",
  "Computers & Peripherals > Computer Storage Devices > Computer HDDs",
  "Computers & Peripherals > Computer Storage Devices > Computer HDDs > 12TB Hard Disk Drives",
  "Computers & Peripherals > Computer Storage Devices > Computer HDDs > 16TB Hard Disk Drives",
  "Computers & Peripherals > Computer Storage Devices > Computer HDDs > 1TB Hard Disk Drives",
  "Computers & Peripherals > Computer Storage Devices > Computer HDDs > 20TB Hard Disk Drives",
  "Computers & Peripherals > Computer Storage Devices > Computer HDDs > 2TB Hard Disk Drives",
  "Computers & Peripherals > Computer Storage Devices > Computer HDDs > 4TB Hard Disk Drives",
  "Computers & Peripherals > Computer Storage Devices > Computer HDDs > 5TB Hard Disk Drives",
  "Computers & Peripherals > Computer Storage Devices > Computer HDDs > External Hard Disk Drives",
  "Computers & Peripherals > Computer Storage Devices > Computer HDDs > Internal Hard Disk Drives",
  "Computers & Peripherals > Computer Storage Devices > Computer SSDs",
  "Computers & Peripherals > Computer Storage Devices > Computer SSDs > 1TB Solid-State Drives",
  "Computers & Peripherals > Computer Storage Devices > Computer SSDs > 2TB Solid-State Drives",
  "Computers & Peripherals > Computer Storage Devices > Computer SSDs > 4TB Solid-State Drives",
  "Computers & Peripherals > Computer Storage Devices > Computer SSDs > 8TB Solid-State Drives",
  "Computers & Peripherals > Desktop Workstations",
  "Computers & Peripherals > Drivers & Storage",
  "Computers & Peripherals > Networking Devices",
  "Computers & Peripherals > Networking Devices > Internet Protocol Cameras",
  "Computers & Peripherals > Networking Devices > Network Interface Controllers",
  "Computers & Peripherals > Networking Devices > Network Switches",
  "Computers & Peripherals > Networking Devices > PoE Injectors",
  "Computers & Peripherals > Networking Devices > SFP Modules",
  "Computers & Peripherals > Networking Devices > Wi-Fi Access Points",
  "Computers & Peripherals > Networking Devices > Wi-Fi Extenders",
  "Computers & Peripherals > Networking Devices > Wi-Fi Routers",
  "Computers & Peripherals > Networking Devices > Wiresless Antennas",
  "Computers & Peripherals > Used Computer Equipment",
  "DJ Equipment > Audio Controllers",
  "DJ Equipment > DJ Accessories",
  "DJ Equipment > DJ Bags",
  "DJ Equipment > DJ Bags & Cases",
  "DJ Equipment > DJ Cases > DJ Controller Cases",
  "DJ Equipment > DJ Cases > DJ Flight Cases",
  "DJ Equipment > DJ Cases > DJ Laptop Cases",
  "DJ Equipment > DJ Cases > DJ Mixer Cases",
  "DJ Equipment > DJ Cases > DJ Rack Cases",
  "DJ Equipment > DJ Cases > DJ Soft Cases",
  "DJ Equipment > DJ Controllers",
  "DJ Equipment > DJ Turntable Accessories",
  "DJ Equipment > DJ Turntables",
  "DJ Equipment > Mixers",
  "DJ Equipment > Other DJ Accessories",
  "DJ Equipment > Turntable Accessories",
  "DJ Equipment > Turntables",
  "DJ Equipment > Used DJ Equipment",
  "Home & Business AV > Business AV > Assisted Listening Devices",
  "Home & Business AV > Business AV > Business AV Accessories",
  "Home & Business AV > Business AV > Commercial Audio Devices",
  "Home & Business AV > Business AV > Commercial Speakers",
  "Home & Business AV > Business AV > Headsets & Intercom Systems",
  "Home & Business AV > Business AV > Security & Surveillance",
  "Home & Business AV > Conferencing Systems > Audio Conferencing Systems",
  "Home & Business AV > Conferencing Systems > Audio Controllers & Mixers",
  "Home & Business AV > Conferencing Systems > Conferencing Accessories",
  "Home & Business AV > Conferencing Systems > Conferencing Cameras",
  "Home & Business AV > Conferencing Systems > Interactive Displays",
  "Home & Business AV > Conferencing Systems > Microphones & Speakerphones",
  "Home & Business AV > Conferencing Systems > Presentation Systems",
  "Home & Business AV > Home AV > CD & Media Players",
  "Home & Business AV > Home AV > Home & Portable Speakers",
  "Home & Business AV > Home AV > Home AV Accessories",
  "Home & Business AV > Home AV > Home Theater Systems",
  "Home & Business AV > Home AV > Receivers & Amplifiers",
  "Home & Business AV > Home AV > Smart Home Devices",
  "Home & Business AV > Other Applications > Cleaning & Sanitization",
  "Home & Business AV > Other Applications > Furniture",
  "Home & Business AV > Other Applications > Installation Tools",
  "Home & Business AV > Other Applications > Outdoor Activity",
  "Home & Business AV > Other Applications > Power Accessories",
  "Home & Business AV > TVs & Displays > Digital Signage Player",
  "Home & Business AV > TVs & Displays > Display Accessories",
  "Home & Business AV > TVs & Displays > TVs & Displays",
  "Microphones > Microphone Accessories > Microphone Adapters",
  "Microphones > Microphone Accessories > Microphone Cases",
  "Microphones > Microphone Accessories > Microphone Clips",
  "Microphones > Microphone Accessories > Microphone Grills",
  "Microphones > Microphone Accessories > Microphone Pop Filters",
  "Microphones > Microphone Accessories > Microphone Shockmounts",
  "Microphones > Microphone Accessories > Microphone Windscreens",
  "Microphones > Used Microphones",
  "Microphones > Wired Microphones > Condenser Microphones",
  "Microphones > Wired Microphones > Dynamic Microphones",
  "Microphones > Wired Microphones > Gooseneck Microphones",
  "Microphones > Wired Microphones > Instrument Microphones",
  "Microphones > Wired Microphones > Lavalier Microphones",
  "Microphones > Wired Microphones > Microphone Accessories",
  "Microphones > Wired Microphones > Podium & Goosenecks",
  "Microphones > Wired Microphones > Ribbon Microphones",
  "Microphones > Wired Microphones > Shotgun Microphones",
  "Microphones > Wired Microphones > USB Microphones",
  "Microphones > Wired Microphones > Vocal Microphones",
  "Microphones > Wireless Microphones > Bodypack Microphone Systems",
  "Microphones > Wireless Microphones > Bodypack Systems",
  "Microphones > Wireless Microphones > Combination Systems",
  "Microphones > Wireless Microphones > Handheld Microphone Systems",
  "Microphones > Wireless Microphones > Handheld Systems",
  "Microphones > Wireless Microphones > Instrument Microphone Systems",
  "Microphones > Wireless Microphones > Instrument Systems",
  "Microphones > Wireless Microphones > Microphones for Wireless Systems",
  "Microphones > Wireless Microphones > Other Wireless Systems",
  "Microphones > Wireless Microphones > Receivers & Transmitters",
  "Microphones > Wireless Microphones > Wireless IEM Systems",
  "Microphones > Wireless Microphones > Wireless Microphone Accessories",
  "Microphones > Wireless Microphones > Wireless Microphone Accessories > Wireless Microphone Receivers",
  "Microphones > Wireless Microphones > Wireless Microphone Accessories > Wireless Microphone Transmitters",
  "Mobile > Cases & Mounts",
  "Mobile > Chargers & Adapters",
  "Mobile > Mobile Audio",
  "Mobile > Mobile Audio Accessories",
  "Mobile > Mobile Cases",
  "Mobile > Mobile Chargers",
  "Mobile > Mobile Mounts",
  "Mobile > Mobile Video",
  "Mobile > Mobile Video Accessories",
  "Mobile > Used Mobile Accessories",
  "Musical Instruments > Amplifiers > Bass Amplifiers",
  "Musical Instruments > Amplifiers > Guitar Amplifiers",
  "Musical Instruments > Amplifiers > Keyboard Amplifiers",
  "Musical Instruments > Amps > Bass Amps",
  "Musical Instruments > Amps > Guitar Amps",
  "Musical Instruments > Amps > Keyboard Amps",
  "Musical Instruments > DI Boxes > Active DI Boxes",
  "Musical Instruments > DI Boxes > Passive DI Boxes",
  "Musical Instruments > Drums > Acoustic Drum Kits",
  "Musical Instruments > Drums > Acoustic Kits",
  "Musical Instruments > Drums > Components & Percussions",
  "Musical Instruments > Drums > Drum Machines",
  "Musical Instruments > Drums > Electronic Drum Kits",
  "Musical Instruments > Drums > Electronic Kits",
  "Musical Instruments > Guitars > Acoustic Guitars",
  "Musical Instruments > Guitars > Bass Guitars",
  "Musical Instruments > Guitars > Classical Guitars",
  "Musical Instruments > Guitars > Classical/Nylon",
  "Musical Instruments > Guitars > Electric Guitars",
  "Musical Instruments > Guitars > Ukuleles",
  "Musical Instruments > Instrument Accessories",
  "Musical Instruments > Instrument Accessories > Amp Accessories",
  "Musical Instruments > Instrument Accessories > Amplifier Accessories > Amp Covers",
  "Musical Instruments > Instrument Accessories > Amplifier Accessories > Amp Racks",
  "Musical Instruments > Instrument Accessories > Amplifier Accessories > Amp Stands",
  "Musical Instruments > Instrument Accessories > Amplifier Accessories > Amp Switchers",
  "Musical Instruments > Instrument Accessories > DI Box",
  "Musical Instruments > Instrument Accessories > Drum Accessories",
  "Musical Instruments > Instrument Accessories > Drum Accessories > Cymbal Stands",
  "Musical Instruments > Instrument Accessories > Drum Accessories > Drum Bags",
  "Musical Instruments > Instrument Accessories > Drum Accessories > Drum Beaters",
  "Musical Instruments > Instrument Accessories > Drum Accessories > Drum Cases",
  "Musical Instruments > Instrument Accessories > Drum Accessories > Drum Double Tom Holders",
  "Musical Instruments > Instrument Accessories > Drum Accessories > Drum Heads",
  "Musical Instruments > Instrument Accessories > Drum Accessories > Drum Hi-Hat Stands",
  "Musical Instruments > Instrument Accessories > Drum Accessories > Drum Pedal Springs",
  "Musical Instruments > Instrument Accessories > Drum Accessories > Drum Pedals",
  "Musical Instruments > Instrument Accessories > Drum Accessories > Drum Snare Wires",
  "Musical Instruments > Instrument Accessories > Drum Accessories > Drum Thrones",
  "Musical Instruments > Instrument Accessories > Drum Accessories > Drum Wing Bolts",
  "Musical Instruments > Instrument Accessories > Guitar Accessories",
  "Musical Instruments > Instrument Accessories > Guitar Accessories > Guitar Bags",
  "Musical Instruments > Instrument Accessories > Guitar Accessories > Guitar Capos",
  "Musical Instruments > Instrument Accessories > Guitar Accessories > Guitar Cases",
  "Musical Instruments > Instrument Accessories > Guitar Accessories > Guitar Footstolls",
  "Musical Instruments > Instrument Accessories > Guitar Accessories > Guitar Hangers",
  "Musical Instruments > Instrument Accessories > Guitar Accessories > Guitar Pickguards",
  "Musical Instruments > Instrument Accessories > Guitar Accessories > Guitar Picks > Extra Heavy Guitar Picks",
  "Musical Instruments > Instrument Accessories > Guitar Accessories > Guitar Picks > Heavy Guitar Picks",
  "Musical Instruments > Instrument Accessories > Guitar Accessories > Guitar Picks > Jazz Guitar Picks",
  "Musical Instruments > Instrument Accessories > Guitar Accessories > Guitar Picks > Medium Guitar Picks",
  "Musical Instruments > Instrument Accessories > Guitar Accessories > Guitar Picks > Thin Guitar Picks",
  "Musical Instruments > Instrument Accessories > Guitar Accessories > Guitar Picks > Triangle Guitar Picks",
  "Musical Instruments > Instrument Accessories > Guitar Accessories > Guitar Pickups",
  "Musical Instruments > Instrument Accessories > Guitar Accessories > Guitar Racks",
  "Musical Instruments > Instrument Accessories > Guitar Accessories > Guitar Slides",
  "Musical Instruments > Instrument Accessories > Guitar Accessories > Guitar Stands",
  "Musical Instruments > Instrument Accessories > Guitar Accessories > Guitar Straps",
  "Musical Instruments > Instrument Accessories > Guitar Accessories > Guitar Strings > 80/20 Bronze Guitar Strings",
  "Musical Instruments > Instrument Accessories > Guitar Accessories > Guitar Strings > Custom Guitar Strings",
  "Musical Instruments > Instrument Accessories > Guitar Accessories > Guitar Strings > Light Guitar Strings",
  "Musical Instruments > Instrument Accessories > Guitar Accessories > Guitar Strings > Medium Guitar Strings",
  "Musical Instruments > Instrument Accessories > Guitar Accessories > Guitar Strings > Nickel-Plated Guitar Strings",
  "Musical Instruments > Instrument Accessories > Guitar Accessories > Guitar Strings > Steel Guitar Strings",
  "Musical Instruments > Instrument Accessories > Instrument Cables",
  "Musical Instruments > Instrument Accessories > Keyboard Accessories",
  "Musical Instruments > Instrument Accessories > Keyboard Accessories > Keyboard Flight Cases",
  "Musical Instruments > Instrument Accessories > Keyboard Accessories > Keyboard Gig Bags",
  "Musical Instruments > Instrument Accessories > Keyboard Accessories > Keyboard Music Stands",
  "Musical Instruments > Instrument Accessories > Keyboard Accessories > Keyboard Pedals",
  "Musical Instruments > Instrument Accessories > Keyboard Accessories > Keyboard Stands",
  "Musical Instruments > Instrument Accessories > Keyboard Accessories > Music Keyboard Covers",
  "Musical Instruments > Instrument Accessories > Keyboard Accessories > Piano Benches",
  "Musical Instruments > Instrument Accessories > Music Stands & Lights",
  "Musical Instruments > Instrument Accessories > Other Instrument Accessories",
  "Musical Instruments > Instrument Accessories > String & Wind Accessories",
  "Musical Instruments > Instrument Accessories > String Instrument Accessories > Cello Bows",
  "Musical Instruments > Instrument Accessories > String Instrument Accessories > String Instrument Cases",
  "Musical Instruments > Instrument Accessories > String Instrument Accessories > String Instrument Strings",
  "Musical Instruments > Instrument Accessories > String Instrument Accessories > String Winders",
  "Musical Instruments > Instrument Accessories > String Instrument Accessories > Violin Bows",
  "Musical Instruments > Instrument Accessories > Tuners & Metronomes",
  "Musical Instruments > Instrument Cables",
  "Musical Instruments > Instrument Effects > Effect Pedal Accessories > Daisy Chain Cables",
  "Musical Instruments > Instrument Effects > Effect Pedal Accessories > Effect Pedal Switches",
  "Musical Instruments > Instrument Effects > Effect Pedal Accessories > Pedal Patch Cables",
  "Musical Instruments > Instrument Effects > Effect Pedal Accessories > Pedal Power Supply",
  "Musical Instruments > Instrument Effects > Effect Pedal Accessories > Pedalboard Cases",
  "Musical Instruments > Instrument Effects > Effect Pedal Accessories > Pedalboards",
  "Musical Instruments > Instrument Effects > Effects Pedals",
  "Musical Instruments > Instrument Effects > Effects Pedals > Distortion Pedals",
  "Musical Instruments > Instrument Effects > Effects Pedals > Filter Pedals",
  "Musical Instruments > Instrument Effects > Effects Pedals > Modulation Effect Pedals",
  "Musical Instruments > Instrument Effects > Effects Pedals > Multi Effect Pedals",
  "Musical Instruments > Instrument Effects > Pedal Accessories",
  "Musical Instruments > Instrument Effects > Rack Effects",
  "Musical Instruments > Keyboard Instruments > Digital Pianos",
  "Musical Instruments > Keyboard Instruments > Midi Interfaces",
  "Musical Instruments > Keyboards & MIDI > Digital Pianos",
  "Musical Instruments > Keyboards & MIDI > MIDI Interfaces",
  "Musical Instruments > Keyboards & MIDI > Workstations & Synthesizers",
  "Musical Instruments > Metronomes",
  "Musical Instruments > Music Stand Lights",
  "Musical Instruments > Music Stands",
  "Musical Instruments > String Instruments > Banjos",
  "Musical Instruments > String Instruments > Mandolins",
  "Musical Instruments > String Instruments > Ukuleles",
  "Musical Instruments > String Instruments > Violins",
  "Musical Instruments > Strings & Winds > Brass & Woodwinds",
  "Musical Instruments > Strings & Winds > Free Reed Instruments",
  "Musical Instruments > Strings & Winds > Strings",
  "Musical Instruments > Tuners",
  "Musical Instruments > Used Muscial Instruments",
  "Musical Instruments > Wind Instruments > Brass Wind Instruments",
  "Musical Instruments > Wind Instruments > Brass Wind Instruments > Clarinets",
  "Musical Instruments > Wind Instruments > Brass Wind Instruments > Saxophones",
  "Musical Instruments > Wind Instruments > Brass Wind Instruments > Trombones",
  "Musical Instruments > Wind Instruments > Brass Wind Instruments > Trumpets",
  "Musical Instruments > Wind Instruments > Electronic Wind Instruments",
  "Musical Instruments > Wind Instruments > Free Reed Instruments",
  "Musical Instruments > Wind Instruments > Woodwind Instruments > Flutes",
  "Musical Instruments > Wind Instruments > Woodwind Instruments > Harmonicas",
  "Pro Audio > Accessories > Audio Utility Carts and Accessories",
  "Pro Audio > Accessories > Microphone Stands",
  "Pro Audio > Accessories > Rack Cases & Accessories",
  "Pro Audio > Accessories > Speaker Stands",
  "Pro Audio > Accessories > Stand Accessories",
  "Pro Audio > Audio Accessories > Carts & Cart Accessories",
  "Pro Audio > Audio Accessories > Microphone Stands",
  "Pro Audio > Audio Accessories > Rack Cases & Accessories",
  "Pro Audio > Audio Accessories > Speaker Stands",
  "Pro Audio > Audio Accessories > Stand Accessories",
  "Pro Audio > Audio Cables > Audio Cables & Cable Accessories",
  "Pro Audio > Audio Cables > Instrument Cables",
  "Pro Audio > Audio Cables > Speakon & 1/4\" Speaker Cables",
  "Pro Audio > Audio Cables > XLR Cables",
  "Pro Audio > Audio Consoles & Mixers > Analog Consoles",
  "Pro Audio > Audio Consoles & Mixers > Audio Mixer Accessories",
  "Pro Audio > Audio Consoles & Mixers > Digital Consoles",
  "Pro Audio > Audio Consoles & Mixers > Powered Mixers",
  "Pro Audio > Audio Consoles & Mixers > Rack Mixers",
  "Pro Audio > Cables > 1/4 Inch Speaker Cables",
  "Pro Audio > Cables > Audio Cable Accessories",
  "Pro Audio > Cables > Instrument Cables",
  "Pro Audio > Cables > XLR Cables",
  "Pro Audio > Consoles > Analog Consoles",
  "Pro Audio > Consoles > Digital Consoles",
  "Pro Audio > Headphones > Headphone Accessories",
  "Pro Audio > Headphones > Headphone Amplifiers",
  "Pro Audio > Headphones > In-Ear Headphones",
  "Pro Audio > Headphones > In-Ear Monitors",
  "Pro Audio > Headphones > On-Ear Headphones",
  "Pro Audio > Headphones > Over-Ear Headphones",
  "Pro Audio > Headphones > Wired Headphones",
  "Pro Audio > Headphones > Wireless Headphones",
  "Pro Audio > Live Sound & Portable Systems > Installed Speakers",
  "Pro Audio > Live Sound & Portable Systems > Karaoke Equipment",
  "Pro Audio > Live Sound & Portable Systems > Kareoke Equipment",
  "Pro Audio > Live Sound & Portable Systems > Live Sound Accessories",
  "Pro Audio > Live Sound & Portable Systems > PA Systems",
  "Pro Audio > Live Sound & Portable Systems > Power Amplifiers",
  "Pro Audio > Live Sound & Portable Systems > Stage Platforms",
  "Pro Audio > Live Sound & Portable Systems > Stage Platforms & Risers",
  "Pro Audio > Live Sound & Portable Systems > Wireless Speakers",
  "Pro Audio > Loudspeakers > Line Arrays",
  "Pro Audio > Loudspeakers > Loudspeakers Line Arrays",
  "Pro Audio > Loudspeakers > Loudspeakers Point Source",
  "Pro Audio > Loudspeakers > Loudspeakers Subwoofers",
  "Pro Audio > Loudspeakers > Point Source",
  "Pro Audio > Loudspeakers > Speaker Accessories",
  "Pro Audio > Loudspeakers > Subwoofers",
  "Pro Audio > Mixers > Audio Mixer Accessories",
  "Pro Audio > Mixers > Powered Mixers",
  "Pro Audio > Mixers > Rack Mixers",
  "Pro Audio > Recording > Acoustic Treatment",
  "Pro Audio > Recording > Audio Control Surfaces",
  "Pro Audio > Recording > Audio Interfaces",
  "Pro Audio > Recording > Audio Recorders",
  "Pro Audio > Recording > Audio Signal Processors",
  "Pro Audio > Recording > Audio Software",
  "Pro Audio > Recording > Control Surfaces",
  "Pro Audio > Recording > Preamplifiers",
  "Pro Audio > Recording > Recording Accessories",
  "Pro Audio > Recording > Signal Processors",
  "Pro Audio > Recording > Studio Monitors",
  "Pro Lighting > Black & UV Lights",
  "Pro Lighting > Cables & Lamps > DMX Cables",
  "Pro Lighting > Cables & Lamps > Other Lighting Cables",
  "Pro Lighting > Cables & Lamps > Replacement Lamps",
  "Pro Lighting > DMX Control > DMX Consoles",
  "Pro Lighting > DMX Control > DMX Software",
  "Pro Lighting > DMX Distribution & Dimmer Packs > Dimmer Units",
  "Pro Lighting > DMX Distribution & Dimmer Packs > DMX Extenders",
  "Pro Lighting > DMX Distribution & Dimmer Packs > DMX Splitters",
  "Pro Lighting > DMX Lighting Equipment > DMX Cables > 3 Pin DMX Cables",
  "Pro Lighting > DMX Lighting Equipment > DMX Cables > 3 Pin to 5 Pin DMX Cables",
  "Pro Lighting > DMX Lighting Equipment > DMX Cables > 5 Pin DMX Cables",
  "Pro Lighting > DMX Lighting Equipment > DMX Cables > 5 Pin to 3 Pin DMX Cables",
  "Pro Lighting > DMX Lighting Equipment > DMX Consoles",
  "Pro Lighting > DMX Lighting Equipment > DMX Software",
  "Pro Lighting > General & Industrial Lighting > Area Lights",
  "Pro Lighting > General & Industrial Lighting > Flashlights",
  "Pro Lighting > General & Industrial Lighting > General Use Lights",
  "Pro Lighting > General & Industrial Lighting > Headlamps",
  "Pro Lighting > Industrial Lighting > General Lights",
  "Pro Lighting > Industrial Lighting > Industrial Headlamps",
  "Pro Lighting > Industrial Lighting > Military Flashlights",
  "Pro Lighting > Industrial Lighting > Remote Area Lights",
  "Pro Lighting > LED Effect Lights",
  "Pro Lighting > LED Fixtures > Bars & Battens",
  "Pro Lighting > LED Fixtures > Black & UV Lights",
  "Pro Lighting > LED Fixtures > Effect Lightings",
  "Pro Lighting > LED Fixtures > Follow Spots",
  "Pro Lighting > LED Fixtures > Moving Heads",
  "Pro Lighting > LED Fixtures > Pars",
  "Pro Lighting > LED Follow Spots",
  "Pro Lighting > LED Lights > LED Stage Bars",
  "Pro Lighting > LED Moving Heads",
  "Pro Lighting > LED Par Lights",
  "Pro Lighting > Lighting Accessories > Barn Doors",
  "Pro Lighting > Lighting Accessories > Controllers & Controller Accessories",
  "Pro Lighting > Lighting Accessories > Light Safety Wires",
  "Pro Lighting > Lighting Accessories > Lighting Barn Doors",
  "Pro Lighting > Lighting Accessories > Lighting Cables & Adapters",
  "Pro Lighting > Lighting Accessories > Lighting Cases",
  "Pro Lighting > Lighting Accessories > Lighting Power Supplies",
  "Pro Lighting > Lighting Accessories > Lighting Stands",
  "Pro Lighting > Lighting Accessories > Other Lighting Accessories",
  "Pro Lighting > Lighting Accessories > Replacement Lamps",
  "Pro Lighting > Lighting Accessories > Safety Wires",
  "Pro Lighting > Lighting Cables",
  "Pro Lighting > Lighting Cables > Replacement Lamps",
  "Pro Lighting > Lighting Cables > Stage Lighting Extension Cords",
  "Pro Lighting > Lighting Trusses > Lighting Rig Clamps",
  "Pro Lighting > Lighting Trusses > Lighting Truss Base Plates",
  "Pro Lighting > Photography Lights > Continuous Lights > Photography LED Light Kits",
  "Pro Lighting > Photography Lights > Continuous Lights > Photography LED Panels",
  "Pro Lighting > Photography Lights > Continuous Lights > Photography Ring Lights",
  "Pro Lighting > Photography Lights > Lighting Modifiers > Photography Reflectors",
  "Pro Lighting > Photography Lights > Lighting Modifiers > Photography Softboxes",
  "Pro Lighting > Photography Lights > Lighting Modifiers > Photography Umbrellas",
  "Pro Lighting > Photography Lights > On Camera LED Lights",
  "Pro Lighting > Photography Lights > Photography Accessories > Photography Backdrops > Collapsible Photography Backdrops",
  "Pro Lighting > Photography Lights > Photography Accessories > Photography Backdrops > Photography Canvas Backdrops",
  "Pro Lighting > Photography Lights > Photography Accessories > Photography Backdrops > Photography Green Screen Backdrops",
  "Pro Lighting > Photography Lights > Photography Accessories > Photography Backdrops > Photography Muslin Backdrops",
  "Pro Lighting > Photography Lights > Photography Accessories > Photography Backdrops > Photography Seamless Backdrops",
  "Pro Lighting > Photography Lights > Photography Accessories > Photography Backdrops > Photography Vinyl Backdrops",
  "Pro Lighting > Photography Lights > Photography Accessories > Photography Lens Filters",
  "Pro Lighting > Production & Photography > Continuous Lighting",
  "Pro Lighting > Production & Photography > Lighting Modifiers",
  "Pro Lighting > Production & Photography > On-Camera Lighting",
  "Pro Lighting > Production & Photography > Photography Accessories",
  "Pro Lighting > Special Effect Machines > Bubble Effect Machines",
  "Pro Lighting > Special Effect Machines > Fog Effect Machines",
  "Pro Lighting > Special Effect Machines > Laser Light Machines",
  "Pro Lighting > Special Effect Machines > Snow Effect Machines",
  "Pro Lighting > Special Effect Machines > Special Effects Machine Accessories > Bubble Machine Fluid",
  "Pro Lighting > Special Effect Machines > Special Effects Machine Accessories > Bubble Machine Remotes",
  "Pro Lighting > Special Effect Machines > Special Effects Machine Accessories > Fog Machine Fagrance",
  "Pro Lighting > Special Effect Machines > Special Effects Machine Accessories > Fog Machine Fluid",
  "Pro Lighting > Special Effect Machines > Special Effects Machine Accessories > Fog Machine Hose Adapters",
  "Pro Lighting > Special Effect Machines > Special Effects Machine Accessories > Fog Machine Remotes",
  "Pro Lighting > Special Effect Machines > Special Effects Machine Accessories > Haze Machine Fluid",
  "Pro Lighting > Special Effect Machines > Special Effects Machine Accessories > Snow Machine Fluid",
  "Pro Lighting > Special FX > Bubble Machines",
  "Pro Lighting > Special FX > Fluids & Accessories",
  "Pro Lighting > Special FX > Fog & Haze Machines",
  "Pro Lighting > Special FX > Lasers Effects",
  "Pro Lighting > Special FX > Smoke Machines",
  "Pro Lighting > Special FX > Snow Machines",
  "Pro Lighting > Truss & Rigging > Rigging Clamps & Accessories",
  "Pro Lighting > Truss & Rigging > Truss",
  "Pro Lighting > Used Lighting Equipment",
  "Pro Video > Camcorder Accessories > Camera Accessories",
  "Pro Video > Camcorder Accessories > Camera Accessories > Camera Bags",
  "Pro Video > Camcorder Accessories > Camera Accessories > Camera Stabilizers",
  "Pro Video > Camcorder Accessories > Camera Backpacks",
  "Pro Video > Camcorder Accessories > Camera Batteries",
  "Pro Video > Camcorder Accessories > Camera Tripods",
  "Pro Video > Camcorder Accessories > Drone Accessories",
  "Pro Video > Camcorder Accessories > Lens Accessories",
  "Pro Video > Camcorders > 4k Camcorders",
  "Pro Video > Camcorders > HD Camcorders",
  "Pro Video > Camera & Camcorder Accessories > Batteries & Battery Accessories",
  "Pro Video > Camera & Camcorder Accessories > Cases & Backpacks",
  "Pro Video > Camera & Camcorder Accessories > Drone Accessories",
  "Pro Video > Camera & Camcorder Accessories > Filters & Filter Accessories",
  "Pro Video > Camera & Camcorder Accessories > Lenses & Lens Accessories",
  "Pro Video > Camera & Camcorder Accessories > Other Camera Accessories",
  "Pro Video > Camera & Camcorder Accessories > Stabilizers & Tripods",
  "Pro Video > Cameras & Camcorders > 4K & 8K Camcorders",
  "Pro Video > Cameras & Camcorders > Cameras",
  "Pro Video > Cameras & Camcorders > Drone & Action Cameras",
  "Pro Video > Cameras & Camcorders > HD Camcorders",
  "Pro Video > Cameras & Camcorders > Mirrorless Cameras",
  "Pro Video > Cameras & Camcorders > PTZ Cameras",
  "Pro Video > Cameras > Action Cameras",
  "Pro Video > Cameras > Camera Accessories",
  "Pro Video > Cameras > Camera Accessories > Camera Backpacks",
  "Pro Video > Cameras > Camera Accessories > Camera Bags",
  "Pro Video > Cameras > Camera Accessories > Camera Batteries",
  "Pro Video > Cameras > Camera Accessories > Camera Stabilizers",
  "Pro Video > Cameras > Camera Accessories > Camera Tripods",
  "Pro Video > Cameras > Compact Cameras",
  "Pro Video > Cameras > Mirrorless Cameras",
  "Pro Video > Cameras > PTZ Cameras",
  "Pro Video > Cameras > Waterproof Cameras",
  "Pro Video > Capture & Recording & Playback > PCIE Recorder",
  "Pro Video > Capture & Recording & Playback > Recorder Accessories",
  "Pro Video > Capture & Recording & Playback > SSD/SD Card Recorder",
  "Pro Video > Capture & Recording & Playback > USB 3.0 & Thunderbolt Recorders",
  "Pro Video > Capture & Recording & Playback > Video Monitors",
  "Pro Video > Computer Video & Software > Computer Hardware",
  "Pro Video > Computer Video & Software > Video Software",
  "Pro Video > Converters & Distribution & Matrix > Distribution Amplifiers",
  "Pro Video > Converters & Distribution & Matrix > Extenders & Repeaters",
  "Pro Video > Converters & Distribution & Matrix > Video Converters",
  "Pro Video > Converters & Distribution & Matrix > Video Matrix",
  "Pro Video > Converters & Distribution & Matrix > Video Processor Accessories",
  "Pro Video > Converters & Distribution & Matrix > Video Splitters",
  "Pro Video > Digital Signage Players",
  "Pro Video > Distribution Amplifiers",
  "Pro Video > Drones",
  "Pro Video > Drones > Drone Accessories",
  "Pro Video > LED Processors",
  "Pro Video > LED Video Wall Accessories",
  "Pro Video > LED Video Wall Accessories > LED Receiving Cards",
  "Pro Video > LED Video Wall Accessories > LED Sending Cards",
  "Pro Video > LED Video Walls > LED Rigging & Accessories",
  "Pro Video > LED Video Walls > LED Video Processors",
  "Pro Video > LED Video Walls > Video Wall Panels",
  "Pro Video > PCI Express Cards",
  "Pro Video > Projectors > DLP Projectors",
  "Pro Video > Projectors > Laser Projectors",
  "Pro Video > Projectors > LCD Projectors",
  "Pro Video > Projectors > Projection Screens",
  "Pro Video > Projectors > Projectors Screens",
  "Pro Video > Switchers & Production Equipment > Live Stream Devices",
  "Pro Video > Switchers & Production Equipment > Teleprompters & Accessories",
  "Pro Video > Switchers & Production Equipment > Video Production Accessories",
  "Pro Video > Switchers & Production Equipment > Video Production Switchers",
  "Pro Video > Switchers & Production Equipment > Video Transmission Systems",
  "Pro Video > TVs > Display Accessories",
  "Pro Video > Used Pro Video",
  "Pro Video > Video Accessories > Memory Cards",
  "Pro Video > Video Accessories > Projector Cases",
  "Pro Video > Video Accessories > Projector Filters",
  "Pro Video > Video Accessories > Projector Lamps",
  "Pro Video > Video Accessories > Projector Lenses",
  "Pro Video > Video Accessories > Projector Mounts",
  "Pro Video > Video Accessories > Projector Remotes",
  "Pro Video > Video Accessories > Video Audio Accessories",
  "Pro Video > Video Cables > HDMI Cable",
  "Pro Video > Video Cables > HDMI Cables",
  "Pro Video > Video Cables > Other Video Cables",
  "Pro Video > Video Cables > SDI Cable",
  "Pro Video > Video Cables > SDI Cables",
  "Pro Video > Video Converters",
  "Pro Video > Video Extenders",
  "Pro Video > Video Hardware",
  "Pro Video > Video Lenses",
  "Pro Video > Video Monitors",
  "Pro Video > Video Monitors > Broadcast Monitors",
  "Pro Video > Video Monitors > Field Monitors",
  "Pro Video > Video Monitors > HDR Monitors",
  "Pro Video > Video Monitors > On-Camera Monitors",
  "Pro Video > Video Monitors > Rackmount Monitors",
  "Pro Video > Video Production Equipment > Live Stream Devices",
  "Pro Video > Video Production Equipment > Video Production Accessories",
  "Pro Video > Video Production Equipment > Video Production Switchers",
  "Pro Video > Video Production Equipment > Video Production Switchers > Broadcast Switchers",
  "Pro Video > Video Production Equipment > Video Production Switchers > Live Production Switchers",
  "Pro Video > Video Production Equipment > Video Production Switchers > Matrix Switchers",
  "Pro Video > Video Production Equipment > Video Production Switchers > Portable Switchers",
  "Pro Video > Video Production Equipment > Video Teleprompters",
  "Pro Video > Video Production Equipment > Video Transmission Systems",
  "Pro Video > Video Recording Accessories",
  "Pro Video > Video Recording Equipment > SDI Card Recorder",
  "Pro Video > Video Software",
  "Pro Video > Video Splitters",
  "Used > Used Computers & Peripherals",
  "Used > Used DJ Equipment",
  "Used > Used Home & Business AV > Used Business AV",
  "Used > Used Home & Business AV > Used Conferencing Systems",
  "Used > Used Home & Business AV > Used Home AV",
  "Used > Used Home & Business AV > Used Other Applications",
  "Used > Used Home & Business AV > Used TVs & Displays",
  "Used > Used Microphones > Used Wired Microphones",
  "Used > Used Microphones > Used Wireless Microphones",
  "Used > Used Mobile > Used iOS & Android",
  "Used > Used Mobile > Used Mobile Accessories",
  "Used > Used Musical Instruments > Used Amps",
  "Used > Used Musical Instruments > Used Drums",
  "Used > Used Musical Instruments > Used Guitars",
  "Used > Used Musical Instruments > Used Instrument Accessories",
  "Used > Used Musical Instruments > Used Instrument Effects",
  "Used > Used Musical Instruments > Used Keyboards & MIDI",
  "Used > Used Musical Instruments > Used Strings & Winds",
  "Used > Used Pro Audio > Used Audio Accessories",
  "Used > Used Pro Audio > Used Audio Cables",
  "Used > Used Pro Audio > Used Audio Consoles & Mixers",
  "Used > Used Pro Audio > Used Headphones",
  "Used > Used Pro Audio > Used Live Sound & Portable Systems",
  "Used > Used Pro Audio > Used Loudspeakers",
  "Used > Used Pro Audio > Used Recording",
  "Used > Used Pro Lighting > Used Cables & Lamps",
  "Used > Used Pro Lighting > Used DMX Control",
  "Used > Used Pro Lighting > Used DMX Distribution & Dimmer Packs",
  "Used > Used Pro Lighting > Used General & Industrial Lighting",
  "Used > Used Pro Lighting > Used Halogen Fixtures",
  "Used > Used Pro Lighting > Used LED Fixtures",
  "Used > Used Pro Lighting > Used Lighting Accessories",
  "Used > Used Pro Lighting > Used Production & Photography",
  "Used > Used Pro Lighting > Used Special FX",
  "Used > Used Pro Lighting > Used Truss & Rigging",
  "Used > Used Pro Video > Used Camera & Camcorder Accessories",
  "Used > Used Pro Video > Used Cameras & Camcorders",
  "Used > Used Pro Video > Used Capture & Recording & Playback",
  "Used > Used Pro Video > Used Computer Video & Software",
  "Used > Used Pro Video > Used Converters & Distribution & Matrix",
  "Used > Used Pro Video > Used LED Video Walls",
  "Used > Used Pro Video > Used Projectors",
  "Used > Used Pro Video > Used Switchers & Production Equipment",
  "Used > Used Pro Video > Used Video Accessories",
  "Used > Used Pro Video > Used Video Cables",
];
// --- END OF USER POPULATION SECTION ---

/**
 * Converts a spreadsheet column letter (e.g., "A" or "FY") to a 1-based column index.
 */
function columnLetterToIndex(col) {
  let index = 0;
  for (let i = 0; i < col.length; i++) {
    index = index * 26 + (col.charCodeAt(i) - 64);
  }
  return index;
}


/**
 * Creates a custom menu when the Google Sheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Product Tools')
    .addItem('1. Set API Key', 'setApiKey') // MODIFIED Text and Function Name
    .addItem('2. Categorize Products (Fixed List)', 'categorizeProductsWithFixedList')
    .addItem('3. Categorize Next Batch', 'categorizeNextBatch')
    .addToUi();
}

/**
 * Prompts the user to set and store their API key.
 */
function setApiKey() { // MODIFIED Function Name
  const ui = SpreadsheetApp.getUi();
  const userProperties = PropertiesService.getUserProperties();
  let result = ui.prompt('API Key Setup', 'Enter your API Key:', ui.ButtonSet.OK_CANCEL); // MODIFIED Text
  if (result.getSelectedButton() === ui.Button.OK && result.getResponseText()) {
    userProperties.setProperty('OPENAI_API_KEY', result.getResponseText().trim()); // Internal property name remains specific
    ui.alert('API Key saved for this user.'); // MODIFIED Text
  } else if (result.getSelectedButton() !== ui.Button.CANCEL) {
    ui.alert('API Key not set or invalid input.'); // MODIFIED Text
  }
}

/**
 * Pre-filters the master category list based on product vendor and keywords.
 */
function getRelevantCategoriesForProduct(productTitle, productDescription, productVendor, allMasterCategories) {
  const titleLower = productTitle ? String(productTitle).toLowerCase() : "";
  const descriptionLower = productDescription ? String(productDescription).toLowerCase() : "";
  const combinedText = titleLower + " " + descriptionLower;
  const vendorLower = productVendor ? String(productVendor).toLowerCase() : "";

  let candidateGroupPrefixes = new Set(); 
  let matchOccurred = false;

  for (const vendorKey in VENDOR_SPECIFIC_CATEGORY_PREFIXES) {
    if (vendorLower.includes(vendorKey.toLowerCase())) { 
      VENDOR_SPECIFIC_CATEGORY_PREFIXES[vendorKey].forEach(prefix => candidateGroupPrefixes.add(prefix));
      matchOccurred = true;
      Logger.log("Matched vendor '" + productVendor + "' to rule for: " + vendorKey);
      break; 
    }
  }

  if (!matchOccurred || candidateGroupPrefixes.size < 5) { 
    KEYWORD_TO_CATEGORY_GROUP_MAPPINGS.forEach(mapping => {
      if (mapping.keywordsRegex.test(combinedText)) {
        mapping.targetGroupPrefixes.forEach(prefix => candidateGroupPrefixes.add(prefix));
        matchOccurred = true; 
      }
    });
  }

  let filteredCategories = [];
  if (candidateGroupPrefixes.size > 0) {
    allMasterCategories.forEach(masterCategory => {
      for (const prefix of candidateGroupPrefixes) {
        if (masterCategory.startsWith(prefix)) {
          filteredCategories.push(masterCategory);
          break; 
        }
      }
    });
    filteredCategories = [...new Set(filteredCategories)]; 
  }

  if (matchOccurred && filteredCategories.length === 0) {
      Logger.log("Warning: Keyword/Vendor rules matched, but no categories found starting with prefixes: [" + Array.from(candidateGroupPrefixes).join(", ") + "] for '" + productTitle + "'. Broadening search.");
      const matchedTopLevels = new Set();
      candidateGroupPrefixes.forEach(prefix => {
          const topLevel = prefix.split(" > ")[0];
          if (TOP_LEVEL_CATEGORIES_FOR_FALLBACK.includes(topLevel)) {
              matchedTopLevels.add(topLevel);
          }
      });
      if (matchedTopLevels.size > 0) {
          allMasterCategories.forEach(masterCategory => {
              for (const topLevel of matchedTopLevels) {
                  if (masterCategory.startsWith(topLevel)) {
                      filteredCategories.push(masterCategory);
                      break;
                  }
              }
          });
          filteredCategories = [...new Set(filteredCategories)];
      }
  }

  if (filteredCategories.length === 0) {
    Logger.log("No categories found from vendor or keyword mapping for '" + productTitle + "'. Using general top-level fallback.");
    allMasterCategories.forEach(masterCategory => {
      for (const topLevel of TOP_LEVEL_CATEGORIES_FOR_FALLBACK) {
        if (masterCategory.startsWith(topLevel)) {
          filteredCategories.push(masterCategory); 
        }
      }
    });
    filteredCategories = [...new Set(filteredCategories)];
  }
  
  if (filteredCategories.length > MAX_CATEGORIES_TO_SEND) {
    Logger.log("Filtered list for '" + productTitle + "' has " + filteredCategories.length + " categories. Truncating to " + MAX_CATEGORIES_TO_SEND);
    filteredCategories = filteredCategories.slice(0, MAX_CATEGORIES_TO_SEND);
  }
  
  if (filteredCategories.length === 0 && allMasterCategories.length > 0) {
      Logger.log("CRITICAL FALLBACK: No categories selected by any filter for '" + productTitle + "'. Sending first 50 master categories to AI to prevent empty list.");
      return allMasterCategories.slice(0, 50); 
  }

  Logger.log("For product '" + productTitle + "' (Vendor: " + productVendor + "), sending " + filteredCategories.length + " relevant categories to AI: [" + filteredCategories.slice(0,5).join(", ") + (filteredCategories.length > 5 ? "..." : "") + "]");
  return filteredCategories;
}


/**
 * Processes product listings to categorize them using a fixed list and an external API.
 */
function categorizeProductsWithFixedList(maxRowsToProcess) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseTemplateSheet = ss.getSheetByName(SHEET_NAME_PRODUCTS);
  const categoriesSheet = ss.getSheetByName(SHEET_NAME_CATEGORIES);

  if (!baseTemplateSheet) { ui.alert(`Error: Sheet "${SHEET_NAME_PRODUCTS}" not found.`); return; }

  let allMasterCategories = [];
  if (categoriesSheet) {
    allMasterCategories = categoriesSheet.getRange("A2:A" + categoriesSheet.getLastRow()).getValues().flat().filter(String);
    if (allMasterCategories.length === 0) {
      Logger.log("No categories found in '" + SHEET_NAME_CATEGORIES + "'. Using built-in list.");
      allMasterCategories = DEFAULT_CATEGORY_LIST;
    }
  } else {
    ui.alert(`Warning: Sheet "${SHEET_NAME_CATEGORIES}" not found. Using built-in category list.`);
    allMasterCategories = DEFAULT_CATEGORY_LIST;
  }

  const apiKey = PropertiesService.getUserProperties().getProperty('OPENAI_API_KEY'); // Internal property name unchanged
  if (!apiKey) {
    ui.alert("Error: API Key not set. Please run '1. Set API Key' from the menu."); // MODIFIED Text
    return;
  }

  Logger.log(`Loaded ${allMasterCategories.length} master categories.`);

  const lastRow = baseTemplateSheet.getLastRow();
  if (lastRow < 2) { ui.alert("No product data found in '" + SHEET_NAME_PRODUCTS + "'."); return; }

  const titleColIndex = columnLetterToIndex(TITLE_COLUMN);
  const descriptionColIndex = columnLetterToIndex(DESCRIPTION_COLUMN);
  const vendorColIndex = columnLetterToIndex(VENDOR_COLUMN);
  const outputColIndex = columnLetterToIndex(OUTPUT_CATEGORY_COLUMN);

  const startCol = Math.min(titleColIndex, descriptionColIndex, vendorColIndex, outputColIndex);
  const endCol = Math.max(titleColIndex, descriptionColIndex, vendorColIndex, outputColIndex);
  const width = endCol - startCol + 1;

  const dataRange = baseTemplateSheet.getRange(2, startCol, lastRow - 1, width);
  const data = dataRange.getValues();

  const titles = [];
  const descriptions = [];
  const vendors = [];
  const currentOutputValues = [];

  for (let i = 0; i < data.length; i++) {
    titles.push(data[i][titleColIndex - startCol]);
    descriptions.push(data[i][descriptionColIndex - startCol]);
    vendors.push(data[i][vendorColIndex - startCol]);
    currentOutputValues.push([data[i][outputColIndex - startCol]]);
  }

  const outputRange = baseTemplateSheet.getRange(2, outputColIndex, lastRow - 1, 1);

  const resultsToWrite = [];
  let processedCount = 0;

  for (let i = 0; i < titles.length; i++) {
    if (maxRowsToProcess && processedCount >= maxRowsToProcess) {
      break;
    }
    if (currentOutputValues[i][0] && !String(currentOutputValues[i][0]).startsWith("API_ERROR") && !String(currentOutputValues[i][0]).startsWith("SCRIPT_ERROR") && String(currentOutputValues[i][0]).toUpperCase() !== "NEEDS_MANUAL_REVIEW") {
      continue;
    }

    const title = titles[i];
    const description = descriptions[i];
    const vendor = vendors[i] || ""; 

    if (!title && !description) {
      continue;
    }

    Logger.log(`Processing row ${i + 2}: Title - "${title}", Vendor - "${vendor}"`);
    processedCount++;
    
    const categoriesForThisPrompt = getRelevantCategoriesForProduct(String(title), String(description), String(vendor), allMasterCategories);

    if (categoriesForThisPrompt.length === 0) {
        Logger.log("CRITICAL: categoriesForThisPrompt is empty for product: " + title + " even after fallbacks. Skipping API call, marking for review.");
        resultsToWrite.push([i, "NEEDS_MANUAL_REVIEW_NO_CATEGORIES_SENT"]);
        continue;
    }

    const categorizationPrompt = `You are an expert e-commerce product categorizer. Carefully study the information below and pick the single most accurate category from the fixed "AVAILABLE CATEGORIES" list.

    Guidelines:
    1. Analyze the Product Title and Description to determine the exact product type, purpose, and any key features.
    2. Consider the vendor/brand for additional context.
    3. Search the list for the most specific category that matches these attributes, using synonyms and context when needed.
    4. If no perfect match exists, choose the closest broader category. Use accessory categories when appropriate.
    5. Ensure your answer is EXACTLY one category name copied verbatim from the list. Do not alter or invent names.

    AVAILABLE CATEGORIES:
    \${categoriesForThisPrompt.join('\\n')}

    Product Title: "\${title}"
    Product Description: "\${description}"

    Think carefully, then reply ONLY with the chosen category:`;

    let chosenCategory = "NEEDS_MANUAL_REVIEW_API_ISSUE"; 
    let attempts = 0;

    while (attempts < MAX_API_RETRIES) {
      const payload = {
        model: OPENAI_MODEL,
        messages: [{ role: "user", content: categorizationPrompt }],
        max_tokens: MAX_COMPLETION_TOKENS,
        temperature: OPENAI_TEMPERATURE
      };
      const options = {
        'method': 'post',
        'headers': { 'Authorization': 'Bearer ' + apiKey, 'Content-Type': 'application/json' },
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': true
      };

      try {
        const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();

        if (responseCode === 200) {
          const jsonResponse = JSON.parse(responseText);
          let extractedCategory = jsonResponse.choices && jsonResponse.choices[0] && jsonResponse.choices[0].message && jsonResponse.choices[0].message.content
                                  ? jsonResponse.choices[0].message.content.trim()
                                  : "";
          
          const lowerCaseAllowedCategories = categoriesForThisPrompt.map(cat => cat.toLowerCase());
          if (extractedCategory && lowerCaseAllowedCategories.includes(extractedCategory.toLowerCase())) {
            chosenCategory = categoriesForThisPrompt.find(cat => cat.toLowerCase() === extractedCategory.toLowerCase()) || extractedCategory;
          } else {
            Logger.log(`CRITICAL AI DEVIATION: AI returned category "${extractedCategory}" which is not in the (filtered) allowed list for title: "${title}". Filtered list had ${categoriesForThisPrompt.length} items. First item was: "${categoriesForThisPrompt.length > 0 ? categoriesForThisPrompt[0] : 'EMPTY_FILTERED_LIST'}". Assigning first from filtered list as fallback.`);
            chosenCategory = categoriesForThisPrompt.length > 0 ? categoriesForThisPrompt[0] : "NEEDS_MANUAL_REVIEW_AI_INVALID_CHOICE";
          }
          break; 

        } else if (responseCode === 429) {
          attempts++;
          Logger.log(`Rate limit hit (attempt ${attempts}/${MAX_API_RETRIES}) for product: ${title}. Retrying after delay. Response: ${responseText}`);
          let retryAfterSeconds = 5 * attempts;
          try {
            const errorPayload = JSON.parse(responseText);
            if (errorPayload && errorPayload.error && errorPayload.error.message) {
              const match = errorPayload.error.message.match(/Please try again in ([\d\.]+)s/);
              if (match && match[1]) {
                retryAfterSeconds = parseFloat(match[1]) + 0.5 + attempts; 
              }
            }
          } catch (e) { /* Ignore parsing error, use calculated retry */ }
          Utilities.sleep(Math.max(retryAfterSeconds * 1000, 2000)); 
          if (attempts >= MAX_API_RETRIES) {
            Logger.log(`Max retry attempts reached for product: ${title}.`);
            chosenCategory = `API_ERROR_RATE_LIMIT_MAX_RETRIES: ${responseCode}`;
            break;
          }
        } else if (responseCode === 401) {
          Logger.log(`Unauthorized API key for product ${title}. Response: ${responseText}`);
          chosenCategory = 'API_ERROR_UNAUTHORIZED';
          break;
        } else {
          Logger.log(`API Error for product ${title}: HTTP ${responseCode} - ${responseText}`);
          chosenCategory = `API_ERROR: ${responseCode}`;
          break;
        }
      } catch (e) {
        Logger.log(`Script execution error during API call for product ${title}: ${e.toString()} ${e.stack}`);
        chosenCategory = `SCRIPT_ERROR: ${e.message}`;
        break; 
      }
    }
    resultsToWrite.push([i, chosenCategory]);
  }

  if (resultsToWrite.length > 0) {
    resultsToWrite.forEach(function(item) {
      outputRange.getCell(item[0] + 1, 1).setValue(item[1]);
    });
    ui.alert(`Product categorization complete! ${processedCount} products were processed/re-processed in this run.`);
  } else {
    ui.alert("No products needed processing in this run or no products found.");
  }
}

function categorizeNextBatch() {
  categorizeProductsWithFixedList(BATCH_SIZE);
}

