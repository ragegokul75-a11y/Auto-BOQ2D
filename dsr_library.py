"""
IS 1200 / CPWD DSR Rate Library and Auto-Mapping Engine for Landscape, Residential & Commercial BOQ.
"""

from typing import Dict, List, Any, Optional
import re

DSR_ITEMS: List[Dict[str, Any]] = [
    # --- LANDSCAPE WORKS ---
    {
        "code": "DSR-L01",
        "category": "Landscape - Softscape & Planting",
        "project_types": ["Landscape", "Residential", "Commercial"],
        "description": "Trenching in ordinary soil up to a depth of 60 cm including removal and stacking of serviceable materials and then disposing of surplus soil, spreading and neatly leveling within a lead of 50 m and making up the trenched area to proper levels by supplying and filling with earth and manure in 3:1 ratio.",
        "unit": "Cu.m",
        "rate": 485.00,
        "calc_mode": "VOLUME",  # Area * Depth
        "default_depth": 0.45,
        "default_width": 1.0,
        "is_ref": "IS 1200 (Part 1) / CPWD Hort.",
        "keywords": ["SOIL", "EARTH", "TRENCH", "PLANT-BED", "SHRUB-BED", "EXCAV-LAND", "SWEET-EARTH", "MANURE"],
    },
    {
        "code": "DSR-L02",
        "category": "Landscape - Softscape & Planting",
        "project_types": ["Landscape", "Residential", "Commercial"],
        "description": "Supplying and stacking of good earth (sweet soil) at site including royalty and carriage up to 5 km lead complete (sludge/weed free agricultural top soil).",
        "unit": "Cu.m",
        "rate": 620.00,
        "calc_mode": "VOLUME",
        "default_depth": 0.30,
        "default_width": 1.0,
        "is_ref": "CPWD Horticulture DSR",
        "keywords": ["GOOD-EARTH", "TOPSOIL", "TOP-SOIL", "FILL-SOIL"],
    },
    {
        "code": "DSR-L03",
        "category": "Landscape - Softscape & Planting",
        "project_types": ["Landscape", "Residential", "Commercial"],
        "description": "Turfing with Selection No. 1 Doob grass / Mexican carpet grass (Zoysia japonica) on prepared bed including dibbling, rolling, watering and maintenance for 30 days till grass forms a thick lawn.",
        "unit": "Sq.m",
        "rate": 295.00,
        "calc_mode": "AREA",
        "default_depth": 1.0,
        "default_width": 1.0,
        "is_ref": "CPWD Horticulture DSR 2.10",
        "keywords": ["LAWN", "GRASS", "TURF", "SOFT-LAWN", "L-LAWN", "LAND-LAWN", "MEADOW", "GREENERY"],
    },
    {
        "code": "DSR-L04",
        "category": "Landscape - Softscape & Planting",
        "project_types": ["Landscape", "Residential", "Commercial"],
        "description": "Planting of dense ornamental shrubs / ground covers (Ficus, Ixora, Hamelia, Plumbago, Wedelia) in prepared shrub beds at 300mm c/c spacing including pit preparation, farmyard manure, neem cake and 30 days maintenance.",
        "unit": "Sq.m",
        "rate": 680.00,
        "calc_mode": "AREA",
        "default_depth": 1.0,
        "default_width": 1.0,
        "is_ref": "CPWD Horticulture Standard",
        "keywords": ["SHRUB", "GROUNDCOVER", "GROUND-COVER", "PLANTER", "PLANTING", "L-SHRUB", "FLOWER-BED"],
    },
    {
        "code": "DSR-L05",
        "category": "Landscape - Softscape & Planting",
        "project_types": ["Landscape", "Residential", "Commercial"],
        "description": "Supplying and planting of specimen Avenue / Shade / Flowering Trees (Plumeria alba, Tabebuia rosea, Peltophorum, Terminalia) of height 2.5m to 3.0m in 90x90x90 cm pits filled with 3:1 good earth and cow dung manure, including bamboo tree guard/staking.",
        "unit": "Nos",
        "rate": 2450.00,
        "calc_mode": "COUNT",
        "default_depth": 1.0,
        "default_width": 1.0,
        "is_ref": "CPWD Horticulture DSR",
        "keywords": ["TREE", "TREES", "L-TREE", "SPECIMEN-TREE", "AVENUE-TREE", "PLUMERIA", "TABEBUIA"],
    },
    {
        "code": "DSR-L06",
        "category": "Landscape - Softscape & Planting",
        "project_types": ["Landscape", "Residential", "Commercial"],
        "description": "Supplying and planting of Royal Palm / Foxtail Palm (Wodyetia bifurcata / Roystonea regia) of clear trunk height 2.0m to 2.5m in 120x120x120 cm pits including soil conditioning, crane/manual handling and guy wire supports.",
        "unit": "Nos",
        "rate": 4800.00,
        "calc_mode": "COUNT",
        "default_depth": 1.0,
        "default_width": 1.0,
        "is_ref": "CPWD Horticulture DSR",
        "keywords": ["PALM", "PALMS", "L-PALM", "FOXTAIL", "ROYAL-PALM", "ARECA"],
    },
    {
        "code": "DSR-L07",
        "category": "Landscape - Softscape & Planting",
        "project_types": ["Landscape", "Residential", "Commercial"],
        "description": "Planting of dense perimeter hedge (Murraya paniculata / Clerodendrum / Bambusa multiplex) in double row at 250mm c/c in 60cm wide x 60cm deep continuous trench.",
        "unit": "Rmt",
        "rate": 420.00,
        "calc_mode": "LENGTH",
        "default_depth": 1.0,
        "default_width": 0.6,
        "is_ref": "CPWD Horticulture DSR",
        "keywords": ["HEDGE", "SCREEN-PLANT", "BAMBOO-SCREEN", "L-HEDGE", "GREEN-BUFFER"],
    },
    {
        "code": "DSR-L08",
        "category": "Landscape - Hardscape & Pathways",
        "project_types": ["Landscape", "Residential", "Commercial"],
        "description": "Providing and laying 60mm thick factory-made cement concrete interlocking paver blocks of M-35 grade made by block making machine with strong vibratory compaction, laid over 50mm thick compacted coarse sand cushioning and joints filled with fine silica sand.",
        "unit": "Sq.m",
        "rate": 1180.00,
        "calc_mode": "AREA",
        "default_depth": 1.0,
        "default_width": 1.0,
        "is_ref": "CPWD DSR 16.68 / IS 15658",
        "keywords": ["PAVER", "PAVING", "WALKWAY", "PATHWAY", "HARDSCAPE", "DRIVEWAY", "L-PAVER", "PLAZA", "DECK", "COURT"],
    },
    {
        "code": "DSR-L09",
        "category": "Landscape - Hardscape & Pathways",
        "project_types": ["Landscape", "Residential", "Commercial"],
        "description": "Providing and laying flamed / shot-blasted Sadarahalli / Cobble / Basalt natural stone flooring 40mm thick over 20mm thick CM 1:4 bed including pointing with polymer modified mortar.",
        "unit": "Sq.m",
        "rate": 2150.00,
        "calc_mode": "AREA",
        "default_depth": 1.0,
        "default_width": 1.0,
        "is_ref": "IS 1200 (Part 11) / CPWD 11.41",
        "keywords": ["COBBLE", "GRANITE-PAVING", "STONE-DECK", "STEPPING-STONE", "NATURAL-STONE", "FEATURE-PLAZA"],
    },
    {
        "code": "DSR-L10",
        "category": "Landscape - Hardscape & Pathways",
        "project_types": ["Landscape", "Residential", "Commercial"],
        "description": "Providing and laying precast cement concrete kerb stone (300x300x150 mm) of M-25 grade in position over 100mm thick PCC 1:3:6 haunching and bedding including jointing in CM 1:3 and painting in two coats.",
        "unit": "Rmt",
        "rate": 640.00,
        "calc_mode": "LENGTH",
        "default_depth": 1.0,
        "default_width": 1.0,
        "is_ref": "CPWD DSR 16.69 / IS 5758",
        "keywords": ["KERB", "CURB", "EDGING", "L-KERB", "EDGE-RESTRAINT", "PATH-EDGE"],
    },
    {
        "code": "DSR-L11",
        "category": "Landscape - Street Furniture & Lighting",
        "project_types": ["Landscape", "Residential", "Commercial"],
        "description": "Supply and installation of architectural outdoor LED Bollard Light (9W-12W, IP65, 3000K Warm White, die-cast aluminum body 800mm height) including PCC foundation pedastal and armoured cable termination.",
        "unit": "Nos",
        "rate": 4650.00,
        "calc_mode": "COUNT",
        "default_depth": 1.0,
        "default_width": 1.0,
        "is_ref": "CPWD Electrical / Landscape Spec",
        "keywords": ["BOLLARD", "LIGHT-BOLLARD", "L-LIGHT", "GARDEN-LIGHT", "SPIKE-LIGHT", "POLE-LIGHT"],
    },
    {
        "code": "DSR-L12",
        "category": "Landscape - Street Furniture & Lighting",
        "project_types": ["Landscape", "Residential", "Commercial"],
        "description": "Supply and fixing of Outdoor Garden Bench (1800mm x 600mm x 450mm) made of seasoned WPC / Teak wood slats on heavy-duty powder-coated cast iron/MS frame anchored to paving with SS expansion bolts.",
        "unit": "Nos",
        "rate": 14500.00,
        "calc_mode": "COUNT",
        "default_depth": 1.0,
        "default_width": 1.0,
        "is_ref": "CPWD Landscape Schedule",
        "keywords": ["BENCH", "SEATING", "L-BENCH", "GARDEN-BENCH", "PERGOLA-SEAT"],
    },
    {
        "code": "DSR-L13",
        "category": "Landscape - Irrigation & Drainage",
        "project_types": ["Landscape", "Residential", "Commercial"],
        "description": "Supplying, laying and jointing uPVC Schedule-40 / Class-3 pressure pipe (32mm/40mm dia) for landscape pop-up & drip irrigation header network in 450mm deep trench including fittings, solvent cement and hydro-testing.",
        "unit": "Rmt",
        "rate": 285.00,
        "calc_mode": "LENGTH",
        "default_depth": 1.0,
        "default_width": 1.0,
        "is_ref": "IS 4985 / CPWD DSR",
        "keywords": ["IRRIGATION", "DRIP", "SPRINKLER-LINE", "PIPE", "L-IRRIG", "WATER-LINE"],
    },
    {
        "code": "DSR-L14",
        "category": "Landscape - Water Features & Structures",
        "project_types": ["Landscape", "Residential", "Commercial"],
        "description": "Construction of Reflective Water Body / Lily Pond including RCC M-30 base slab & retaining walls, crystalline integral waterproofing, glass mosaic tile lining and coping stone complete.",
        "unit": "Sq.m",
        "rate": 5400.00,
        "calc_mode": "AREA",
        "default_depth": 1.0,
        "default_width": 1.0,
        "is_ref": "IS 1200 / CPWD Civil + Finishes",
        "keywords": ["WATER-BODY", "POND", "FOUNTAIN", "POOL", "CASCADE", "L-WATER"],
    },

    # --- RESIDENTIAL & COMMERCIAL CIVIL / STRUCTURAL ---
    {
        "code": "DSR-C01",
        "category": "Earthwork & Foundation",
        "project_types": ["Residential", "Commercial", "Landscape"],
        "description": "Earth work in excavation by mechanical means (Hydraulic excavator) / manual means over areas and foundation trenches in all kinds of soil, Depth up to 1.5 m, including dressing of sides and ramming of bottoms, lift up to 1.5 m, including getting out the excavated soil and disposal within 50m lead.",
        "unit": "Cu.m",
        "rate": 265.00,
        "calc_mode": "VOLUME",
        "default_depth": 1.50,
        "default_width": 1.0,
        "is_ref": "IS 1200 (Part 1) / CPWD DSR 2.6.1",
        "keywords": ["EXCAVATION", "EXCAV", "TRENCH-EXCAV", "PIT", "FOUNDATION-PIT"],
    },
    {
        "code": "DSR-C02",
        "category": "Earthwork & Foundation",
        "project_types": ["Residential", "Commercial", "Landscape"],
        "description": "Providing and laying in position cement concrete of specified grade excluding the cost of centering and shuttering - All work up to plinth level: 1:4:8 (1 Cement : 4 coarse sand : 8 graded stone aggregate 40 mm nominal size) for PCC leveling course.",
        "unit": "Cu.m",
        "rate": 6350.00,
        "calc_mode": "VOLUME",
        "default_depth": 0.10,
        "default_width": 1.0,
        "is_ref": "IS 1200 (Part 2) / CPWD DSR 4.1.8",
        "keywords": ["PCC", "BED-CONCRETE", "LEAN-CONCRETE", "SOLING"],
    },
    {
        "code": "DSR-C03",
        "category": "RCC & Structural Works",
        "project_types": ["Residential", "Commercial"],
        "description": "Providing and laying in position machine batched and machine mixed design mix M-25 grade reinforced cement concrete (RCC) for footings, pedestals, columns, plinth beams and floor slabs, including pumping of concrete to site of laying, centering/shuttering and reinforcement measured separately or composite.",
        "unit": "Cu.m",
        "rate": 9850.00,
        "calc_mode": "VOLUME",
        "default_depth": 0.15,
        "default_width": 1.0,
        "is_ref": "IS 1200 (Part 2) / CPWD DSR 5.33",
        "keywords": ["RCC", "SLAB", "FOOTING", "FOUNDATION", "COLUMN-CONC", "BEAM", "PLINTH-BEAM", "STRUCT-SLAB"],
    },
    {
        "code": "DSR-C04",
        "category": "RCC & Structural Works",
        "project_types": ["Residential", "Commercial"],
        "description": "RCC Structural Column / Pier (300x450mm / 300x600mm average section, 3.2m floor height) in M-30 grade concrete complete with Fe-500D TMT reinforcement @ 180 kg/Cu.m and marine ply formwork.",
        "unit": "Nos",
        "rate": 16800.00,
        "calc_mode": "COUNT",
        "default_depth": 3.20,
        "default_width": 0.45,
        "is_ref": "IS 1200 (Part 2 & 5) / CPWD DSR",
        "keywords": ["COLUMN", "COL", "PILLAR", "S-COLS", "STRUCT-COL"],
    },
    {
        "code": "DSR-C05",
        "category": "Masonry & Partition Walls",
        "project_types": ["Residential", "Commercial", "Landscape"],
        "description": "Brick work with common burnt clay F.P.S. (non-modular) bricks of class designation 7.5 / Autoclaved Aerated Concrete (AAC) blocks of 230mm thickness in superstructure above plinth level up to floor V level in cement mortar 1:6 (1 cement : 6 coarse sand).",
        "unit": "Cu.m",
        "rate": 7250.00,
        "calc_mode": "WALL_VOLUME",  # Length * Width(0.23) * Height(3.0)
        "default_depth": 3.00,
        "default_width": 0.23,
        "is_ref": "IS 1200 (Part 3) / CPWD DSR 6.4.2",
        "keywords": ["WALL", "WALL-EXT", "BRICKWORK", "MASONRY", "230MM-WALL", "A-WALL", "EXT-WALL", "COMPOUND-WALL"],
    },
    {
        "code": "DSR-C06",
        "category": "Masonry & Partition Walls",
        "project_types": ["Residential", "Commercial"],
        "description": "Half brick masonry (115mm thick) with common burnt clay F.P.S. bricks / 100mm AAC block partition wall in superstructure in cement mortar 1:4 with 2 Nos. 6mm dia M.S. bars at every third course.",
        "unit": "Sq.m",
        "rate": 1040.00,
        "calc_mode": "WALL_AREA",  # Length * Height(3.0)
        "default_depth": 3.00,
        "default_width": 0.115,
        "is_ref": "IS 1200 (Part 3) / CPWD DSR 6.13.2",
        "keywords": ["PARTITION", "WALL-INT", "115MM-WALL", "HALF-BRICK", "INT-WALL", "TOILET-WALL"],
    },
    {
        "code": "DSR-C07",
        "category": "Commercial Glazing & Partitions",
        "project_types": ["Commercial", "Residential"],
        "description": "Providing and fixing frameless / slim-profile anodized aluminum glazed office partition with 12mm thick toughened clear float glass including patch fittings, silicone sealant and acoustic gasket up to 3.0m height.",
        "unit": "Sq.m",
        "rate": 4250.00,
        "calc_mode": "WALL_AREA",
        "default_depth": 3.00,
        "default_width": 0.05,
        "is_ref": "IS 1200 (Part 14) / CPWD DSR 21.1",
        "keywords": ["GLAZING", "GLASS-PARTITION", "CURTAIN-WALL", "FACADE", "CABIN-GLASS", "A-GLAZ"],
    },

    # --- FLOORING, PLASTERING & ARCHITECTURAL FINISHES ---
    {
        "code": "DSR-F01",
        "category": "Flooring & Tiling",
        "project_types": ["Residential", "Commercial"],
        "description": "Providing and laying Vitrified floor tiles (Double Charged / GVT, 800x800mm or 600x1200mm) in different sizes with water absorption less than 0.08% and conforming to IS:15622, over 20mm thick bed of cement mortar 1:4 and jointed with epoxy grout.",
        "unit": "Sq.m",
        "rate": 1680.00,
        "calc_mode": "AREA",
        "default_depth": 1.0,
        "default_width": 1.0,
        "is_ref": "IS 1200 (Part 11) / CPWD DSR 11.41A",
        "keywords": ["FLOOR", "FLOORING", "TILE", "VITRIFIED", "LIVING-FLOOR", "BEDROOM-FLOOR", "OFFICE-FLOOR", "LOBBY-FLOOR", "A-FLOR", "ROOM"],
    },
    {
        "code": "DSR-F02",
        "category": "Flooring & Tiling",
        "project_types": ["Residential", "Commercial"],
        "description": "Providing and laying Anti-skid Ceramic / Matt Vitrified floor tiles (300x300mm or 600x600mm) in toilets, balconies and wet utility areas over 20mm CM 1:4 bed with waterproof joint grouting.",
        "unit": "Sq.m",
        "rate": 1240.00,
        "calc_mode": "AREA",
        "default_depth": 1.0,
        "default_width": 1.0,
        "is_ref": "IS 1200 (Part 11) / CPWD DSR 11.37",
        "keywords": ["TOILET-FLOOR", "WET-AREA", "BALCONY-FLOOR", "UTILITY-FLOOR", "BATHROOM", "WASHROOM", "RESTROOM"],
    },
    {
        "code": "DSR-F03",
        "category": "Flooring & Tiling",
        "project_types": ["Residential", "Commercial"],
        "description": "Providing and laying 18mm thick mirror-polished Granite stone slab (Jet Black / Sadarahalli / Lakha Red) in entrance lobby, staircase treads & risers, and lift fascia over 20mm CM 1:4 bed including edge moulding and polishing.",
        "unit": "Sq.m",
        "rate": 2890.00,
        "calc_mode": "AREA",
        "default_depth": 1.0,
        "default_width": 1.0,
        "is_ref": "IS 1200 (Part 11) / CPWD DSR 8.2",
        "keywords": ["GRANITE", "MARBLE", "STAIR-FINISH", "CORRIDOR-FLOOR", "LIFT-LOBBY", "COUNTER"],
    },
    {
        "code": "DSR-F04",
        "category": "Plastering, Painting & Ceiling",
        "project_types": ["Residential", "Commercial", "Landscape"],
        "description": "12mm / 15mm cement plaster of mix 1:4 (1 cement : 4 fine sand) on internal & external masonry surfaces including 2 coats of white cement putty and 2 coats of Premium Acrylic Emulsion / Exterior Weather-shield paint.",
        "unit": "Sq.m",
        "rate": 580.00,
        "calc_mode": "WALL_AREA",
        "default_depth": 3.00,
        "default_width": 1.0,
        "is_ref": "IS 1200 (Part 12 & 13) / CPWD DSR 13.1 + 13.60",
        "keywords": ["PLASTER", "PAINT", "WALL-FINISH", "CLADDING"],
    },
    {
        "code": "DSR-F05",
        "category": "Plastering, Painting & Ceiling",
        "project_types": ["Residential", "Commercial"],
        "description": "Providing and fixing suspended Gypsum Board / Acoustic Mineral Fiber Grid False Ceiling (12.5mm tapered edge Gypboard on GI perimeter channels, ceiling angles and intermediate channels at 450mm c/c) including shadow gap, joint taping, primer and 2 coats of luxury emulsion paint.",
        "unit": "Sq.m",
        "rate": 1150.00,
        "calc_mode": "AREA",
        "default_depth": 1.0,
        "default_width": 1.0,
        "is_ref": "IS 1200 / CPWD DSR 12.45",
        "keywords": ["CEILING", "FALSE-CEILING", "GYPSUM", "GRID-CEILING", "RCP", "A-CEIL"],
    },

    # --- DOORS, WINDOWS, MEP & FIXTURES ---
    {
        "code": "DSR-D01",
        "category": "Doors, Windows & Joinery",
        "project_types": ["Residential", "Commercial"],
        "description": "Providing and fixing 35mm thick solid core Flush Door Shutter (1000x2100mm) finished with 1.0mm decorative laminate on both sides in 2nd class Teak Wood / WPC frame (100x75mm) complete with SS mortise lock, ball-bearing hinges, door stopper and tower bolts.",
        "unit": "Nos",
        "rate": 16500.00,
        "calc_mode": "COUNT",
        "default_depth": 1.0,
        "default_width": 1.0,
        "is_ref": "IS 1200 (Part 21) / CPWD DSR 9.21",
        "keywords": ["DOOR", "DOORS", "D1", "D2", "MAIN-DOOR", "FLUSH-DOOR", "A-DOOR", "FIRE-DOOR"],
    },
    {
        "code": "DSR-D02",
        "category": "Doors, Windows & Joinery",
        "project_types": ["Residential", "Commercial"],
        "description": "Providing and fixing UPVC / Powder-coated Aluminum 3-track sliding window with mosquito mesh shutter (average size 1500x1500mm) glazed with 6mm toughened / DGU clear glass including EPDM gaskets, multi-point locking and wool pile.",
        "unit": "Nos",
        "rate": 14200.00,
        "calc_mode": "COUNT",
        "default_depth": 1.0,
        "default_width": 1.0,
        "is_ref": "IS 1200 / CPWD DSR 9.147",
        "keywords": ["WINDOW", "WINDOWS", "W1", "W2", "VENTILATOR", "A-WIND", "SLIDING-WINDOW"],
    },
    {
        "code": "DSR-M01",
        "category": "MEP, Sanitary & HVAC Fixtures",
        "project_types": ["Residential", "Commercial"],
        "description": "Supply and installation of Wall-hung European Water Closet (EWC) with concealed cistern, dual flush actuator plate, soft-close seat cover, CP health faucet and angle valve (Jaquar / Hindware / Kohler equivalent).",
        "unit": "Nos",
        "rate": 18500.00,
        "calc_mode": "COUNT",
        "default_depth": 1.0,
        "default_width": 1.0,
        "is_ref": "CPWD Sanitary Schedule 17.2",
        "keywords": ["WC", "EWC", "TOILET-FIXTURE", "SANITARY", "WASH-BASIN", "BASIN", "URINAL", "PLUMBING-FIX"],
    },
    {
        "code": "DSR-M02",
        "category": "MEP, Sanitary & HVAC Fixtures",
        "project_types": ["Commercial", "Residential"],
        "description": "Supply, installation, testing and commissioning of 2.0 TR 4-Way Inverter Cassette / Split Air Conditioning Indoor Unit with copper refrigerant piping, nitrile insulation, drain piping and wired thermostat.",
        "unit": "Nos",
        "rate": 68000.00,
        "calc_mode": "COUNT",
        "default_depth": 1.0,
        "default_width": 1.0,
        "is_ref": "CPWD HVAC Schedule",
        "keywords": ["AC", "HVAC", "CASSETTE-AC", "SPLIT-AC", "DIFFUSER", "AHU", "VRF"],
    },
    {
        "code": "DSR-M03",
        "category": "MEP, Sanitary & HVAC Fixtures",
        "project_types": ["Commercial", "Residential"],
        "description": "Supply and installation of Recessed LED Panel / Downlight Fixture (18W/36W 600x600mm or 200mm dia, 4000K Neutral White) including FRLS 3x1.5 sq.mm copper wiring in HMS PVC conduit.",
        "unit": "Nos",
        "rate": 2450.00,
        "calc_mode": "COUNT",
        "default_depth": 1.0,
        "default_width": 1.0,
        "is_ref": "CPWD Electrical DSR 1.24",
        "keywords": ["LIGHT", "DOWNLIGHT", "LED-PANEL", "CHANDELIER", "E-LIGHT", "CEILING-LIGHT"],
    },
    {
        "code": "DSR-M04",
        "category": "MEP, Sanitary & HVAC Fixtures",
        "project_types": ["Commercial", "Residential"],
        "description": "Supply and installation of Modular Office Workstation / Executive Desk (1500x750x750mm) made of 25mm pre-laminated E1 grade particle board top with 2mm PVC edge banding, powder-coated MS understructure, wire manager and 3-drawer pedestal.",
        "unit": "Nos",
        "rate": 15800.00,
        "calc_mode": "COUNT",
        "default_depth": 1.0,
        "default_width": 1.0,
        "is_ref": "CPWD Interior & Furniture DSR",
        "keywords": ["WORKSTATION", "DESK", "CHAIR", "FURNITURE", "CUBICLE", "CONFERENCE-TABLE", "SOFA", "BED"],
    },
]

ITEM_BY_CODE: Dict[str, Dict[str, Any]] = {item["code"]: item for item in DSR_ITEMS}


def auto_map_layer_or_block(
    name: str,
    entity_kind: str,  # "AREA", "LENGTH", "COUNT"
    project_type: str = "All",
) -> Dict[str, Any]:
    """
    Intelligently matches a CAD layer name or block name to the best IS 1200 / CPWD DSR item.
    """
    clean_name = name.upper().replace("_", "-").replace(" ", "-")

    best_item: Optional[Dict[str, Any]] = None
    best_score = -1

    for item in DSR_ITEMS:
        score = 0
        # Check keyword matches
        for kw in item["keywords"]:
            if kw == clean_name:
                score += 100
            elif kw in clean_name or clean_name in kw:
                score += 45
            else:
                # Token match
                tokens = re.split(r"[-_./]", clean_name)
                if kw in tokens:
                    score += 65

        # Geometry compatibility bonus
        calc_mode = item["calc_mode"]
        if entity_kind == "COUNT" and calc_mode == "COUNT":
            score += 30
        elif entity_kind == "AREA" and calc_mode in ("AREA", "VOLUME"):
            score += 30
        elif entity_kind == "LENGTH" and calc_mode in ("LENGTH", "WALL_VOLUME", "WALL_AREA", "VOLUME"):
            score += 30
        elif entity_kind == "COUNT" and calc_mode != "COUNT":
            score -= 50  # Blocks should map to COUNT items

        # Project type preference
        if project_type != "All" and project_type in item["project_types"]:
            score += 10

        if score > best_score:
            best_score = score
            best_item = item

    # Fallbacks by entity kind if no keyword matched
    if best_item is None or best_score < 20:
        if entity_kind == "COUNT":
            best_item = ITEM_BY_CODE["DSR-L05" if project_type == "Landscape" else "DSR-D01"]
        elif entity_kind == "AREA":
            best_item = ITEM_BY_CODE["DSR-L03" if project_type == "Landscape" else "DSR-F01"]
        else:
            best_item = ITEM_BY_CODE["DSR-L10" if project_type == "Landscape" else "DSR-C05"]

    return best_item
