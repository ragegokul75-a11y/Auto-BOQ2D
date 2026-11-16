"""
Generates 3 rich, realistic 2D AutoCAD (.DXF) sample drawings in `samples/`:
1. `sample_landscape_masterplan.dxf` — Landscape Masterplan (Lawns, Shrub beds, Paver walkways, Granite plaza, Water pond, Kerb stones, Hedges, Irrigation header, Avenue Trees, Royal Palms, Bollard Lights, Garden Benches).
2. `sample_residential_villa.dxf` — 3BHK Residential Villa Plan (230mm External Brickwork, 115mm Internal Partitions, Vitrified Flooring, Anti-skid Toilet/Balcony tiles, Granite Staircase, RCC Columns, Flush Doors, UPVC Windows, Sanitary EWC).
3. `sample_commercial_office.dxf` — Commercial Corporate Office Floorplate (Glazed Cabin Partitions, Open Workstation Flooring, Acoustic Gypsum False Ceiling, Granite Lift Lobby, Cassette ACs, LED Panel Lights, Modular Workstations, Fire Doors).
"""

import os
import ezdxf

SAMPLES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "samples")


def _add_block_symbol(doc: ezdxf.document.Drawing, name: str, radius: float = 0.6, color: int = 3):
    if name in doc.blocks:
        return
    blk = doc.blocks.new(name=name)
    blk.add_circle((0, 0), radius=radius, dxfattribs={"color": color})
    blk.add_line((-radius * 0.7, 0), (radius * 0.7, 0), dxfattribs={"color": color})
    blk.add_line((0, -radius * 0.7), (0, radius * 0.7), dxfattribs={"color": color})


def create_landscape_sample(filepath: str) -> str:
    doc = ezdxf.new("R2010")
    doc.header["$INSUNITS"] = 6  # Meters
    msp = doc.modelspace()

    layers = [
        ("L-LAWN-GRASS", 3),       # Green
        ("L-SHRUB-BED", 80),       # Emerald
        ("L-PAVER-WALKWAY", 30),   # Orange
        ("L-NATURAL-STONE-DECK", 40),
        ("L-WATER-BODY", 4),       # Cyan
        ("L-KERB-EDGING", 2),      # Yellow
        ("L-HEDGE-SCREEN", 90),    # Teal
        ("L-IRRIGATION-PIPE", 5),  # Blue
        ("L-TREES", 3),
        ("L-PALMS", 80),
        ("L-LIGHTING", 2),
        ("L-FURNITURE", 30),
        ("ANNO", 7),
    ]
    for lname, col in layers:
        if lname not in doc.layers:
            doc.layers.add(lname, color=col)

    _add_block_symbol(doc, "SPECIMEN_TREE", radius=1.4, color=3)
    _add_block_symbol(doc, "ROYAL_PALM", radius=1.1, color=80)
    _add_block_symbol(doc, "BOLLARD_LIGHT", radius=0.35, color=2)
    _add_block_symbol(doc, "GARDEN_BENCH", radius=0.55, color=30)

    # 1. Central Lawn Zones (Closed Polylines)
    msp.add_lwpolyline(
        [(4, 6), (26, 6), (26, 22), (4, 22)],
        close=True,
        dxfattribs={"layer": "L-LAWN-GRASS"},
    )
    msp.add_text("CENTRAL PARTY LAWN (352 Sq.m)", dxfattribs={"layer": "ANNO", "height": 0.7}).set_placement((10, 14))

    msp.add_lwpolyline(
        [(30, 6), (48, 6), (48, 18), (30, 18)],
        close=True,
        dxfattribs={"layer": "L-LAWN-GRASS"},
    )
    msp.add_text("EAST MEADOW LAWN", dxfattribs={"layer": "ANNO", "height": 0.7}).set_placement((35, 12))

    # 2. Ornamental Shrub Beds
    msp.add_lwpolyline(
        [(2, 2), (50, 2), (50, 4.5), (2, 4.5)],
        close=True,
        dxfattribs={"layer": "L-SHRUB-BED"},
    )
    msp.add_text("SOUTH PERIMETER SHRUB BED", dxfattribs={"layer": "ANNO", "height": 0.55}).set_placement((18, 3.0))

    msp.add_lwpolyline(
        [(4, 25), (24, 25), (24, 29), (4, 29)],
        close=True,
        dxfattribs={"layer": "L-SHRUB-BED"},
    )
    msp.add_text("NORTH FLOWERING SHRUB BUFFER", dxfattribs={"layer": "ANNO", "height": 0.55}).set_placement((8, 27))

    # 3. Interlocking Paver Jogging Track & Walkways
    msp.add_lwpolyline(
        [(26.5, 5), (29.5, 5), (29.5, 29), (26.5, 29)],
        close=True,
        dxfattribs={"layer": "L-PAVER-WALKWAY"},
    )
    msp.add_text("SPINE PAVER PROMENADE (3.0m WIDE)", dxfattribs={"layer": "ANNO", "height": 0.5}).set_placement((26.7, 17))

    msp.add_lwpolyline(
        [(4, 22.2), (26.5, 22.2), (26.5, 24.7), (4, 24.7)],
        close=True,
        dxfattribs={"layer": "L-PAVER-WALKWAY"},
    )
    msp.add_text("UPPER JOGGING WALKWAY", dxfattribs={"layer": "ANNO", "height": 0.5}).set_placement((11, 23.2))

    # 4. Natural Stone Plaza Deck
    msp.add_lwpolyline(
        [(30, 19), (41, 19), (41, 29), (30, 29)],
        close=True,
        dxfattribs={"layer": "L-NATURAL-STONE-DECK"},
    )
    msp.add_text("GRANITE / COBBLE VIEWING PLAZA", dxfattribs={"layer": "ANNO", "height": 0.6}).set_placement((31.2, 24))

    # 5. Reflective Lily Pond / Water Body
    msp.add_lwpolyline(
        [(42, 19), (49, 19), (49, 29), (42, 29)],
        close=True,
        dxfattribs={"layer": "L-WATER-BODY"},
    )
    msp.add_text("REFLECTIVE LILY POND", dxfattribs={"layer": "ANNO", "height": 0.55}).set_placement((42.8, 24))

    # 6. Precast Kerb Stone Edging (Linear runs)
    msp.add_lwpolyline(
        [(4, 22.2), (26.5, 22.2), (26.5, 5)],
        close=False,
        dxfattribs={"layer": "L-KERB-EDGING"},
    )
    msp.add_lwpolyline(
        [(29.5, 5), (29.5, 29), (4, 29)],
        close=False,
        dxfattribs={"layer": "L-KERB-EDGING"},
    )

    # 7. Perimeter Green Hedge Screen
    msp.add_lwpolyline(
        [(1.5, 1.5), (50.5, 1.5), (50.5, 30.5), (1.5, 30.5)],
        close=False,
        dxfattribs={"layer": "L-HEDGE-SCREEN"},
    )

    # 8. uPVC Irrigation Mainline Network
    msp.add_lwpolyline(
        [(3, 3), (47, 3), (47, 27), (6, 27)],
        close=False,
        dxfattribs={"layer": "L-IRRIGATION-PIPE"},
    )

    # 9. Block Inserts: Specimen Avenue Trees (12 Nos)
    tree_coords = [
        (6, 8), (12, 8), (18, 8), (24, 8),
        (6, 20), (12, 20), (18, 20), (24, 20),
        (33, 8), (39, 8), (45, 8), (45, 15),
    ]
    for tx, ty in tree_coords:
        msp.add_blockref("SPECIMEN_TREE", (tx, ty), dxfattribs={"layer": "L-TREES"})

    # 10. Block Inserts: Royal Palms along Spine Promenade (8 Nos)
    palm_coords = [
        (25.5, 7), (25.5, 11), (25.5, 15), (25.5, 19),
        (30.5, 7), (30.5, 11), (30.5, 15), (30.5, 19),
    ]
    for px, py in palm_coords:
        msp.add_blockref("ROYAL_PALM", (px, py), dxfattribs={"layer": "L-PALMS"})

    # 11. Block Inserts: LED Bollard Lights (14 Nos)
    light_coords = [
        (5, 22.5), (10, 22.5), (15, 22.5), (20, 22.5), (25, 22.5),
        (26.8, 7), (26.8, 12), (26.8, 17), (26.8, 22), (26.8, 27),
        (32, 19.5), (36, 19.5), (40, 19.5), (44, 19.5),
    ]
    for lx, ly in light_coords:
        msp.add_blockref("BOLLARD_LIGHT", (lx, ly), dxfattribs={"layer": "L-LIGHTING"})

    # 12. Block Inserts: Outdoor Garden Benches (6 Nos)
    bench_coords = [
        (8, 24.2), (16, 24.2), (23, 24.2),
        (32, 27.5), (36, 27.5), (40, 27.5),
    ]
    for bx, by in bench_coords:
        msp.add_blockref("GARDEN_BENCH", (bx, by), dxfattribs={"layer": "L-FURNITURE"})

    doc.saveas(filepath)
    return filepath


def create_residential_sample(filepath: str) -> str:
    doc = ezdxf.new("R2010")
    doc.header["$INSUNITS"] = 6  # Meters
    msp = doc.modelspace()

    layers = [
        ("A-WALL-230MM", 1),          # Red - External 230mm Brick Wall
        ("A-PARTITION-115MM", 30),    # Orange - Internal 115mm Partition Wall
        ("A-FLOOR-VITRIFIED", 4),     # Cyan - Living/Bedroom Vitrified Flooring
        ("A-TOILET-FLOOR", 5),        # Blue - Anti-skid Toilet/Balcony Flooring
        ("A-GRANITE-STAIR", 6),       # Magenta - Granite Lobby & Staircase
        ("S-RCC-SLAB", 8),            # Grey - Plinth / Roof RCC Slab
        ("S-COLUMN", 2),              # Yellow - RCC Columns
        ("A-DOOR", 3),                # Green - Flush Doors
        ("A-WINDOW", 40),             # Amber - UPVC Sliding Windows
        ("M-SANITARY-WC", 5),         # Sanitary EWC Fixtures
        ("ANNO", 7),
    ]
    for lname, col in layers:
        if lname not in doc.layers:
            doc.layers.add(lname, color=col)

    _add_block_symbol(doc, "RCC_COLUMN", radius=0.25, color=2)
    _add_block_symbol(doc, "FLUSH_DOOR", radius=0.45, color=3)
    _add_block_symbol(doc, "UPVC_WINDOW", radius=0.40, color=40)
    _add_block_symbol(doc, "EWC_SANITARY", radius=0.30, color=5)

    # Outer 230mm Brick Wall Perimeter (16m x 12m Villa)
    msp.add_lwpolyline(
        [(0, 0), (16, 0), (16, 12), (0, 12), (0, 0)],
        close=True,
        dxfattribs={"layer": "A-WALL-230MM"},
    )
    # Central Load-Bearing Spine Wall
    msp.add_line((0, 6.5), (16, 6.5), dxfattribs={"layer": "A-WALL-230MM"})

    # Internal 115mm Partition Walls
    msp.add_line((6.5, 0), (6.5, 6.5), dxfattribs={"layer": "A-PARTITION-115MM"})
    msp.add_line((11.5, 0), (11.5, 6.5), dxfattribs={"layer": "A-PARTITION-115MM"})
    msp.add_line((5.5, 6.5), (5.5, 12), dxfattribs={"layer": "A-PARTITION-115MM"})
    msp.add_line((8.5, 6.5), (8.5, 12), dxfattribs={"layer": "A-PARTITION-115MM"})
    msp.add_line((13.0, 6.5), (13.0, 12), dxfattribs={"layer": "A-PARTITION-115MM"})

    # Room Flooring Polygons (Vitrified Tiles)
    # Living & Dining Hall (6.3m x 6.3m)
    msp.add_lwpolyline(
        [(0.2, 0.2), (6.4, 0.2), (6.4, 6.4), (0.2, 6.4)],
        close=True,
        dxfattribs={"layer": "A-FLOOR-VITRIFIED"},
    )
    msp.add_text("LIVING & DINING (6.2x6.2m)", dxfattribs={"layer": "ANNO", "height": 0.35}).set_placement((1.2, 3.2))

    # Kitchen (4.8m x 6.2m)
    msp.add_lwpolyline(
        [(6.6, 0.2), (11.4, 0.2), (11.4, 6.4), (6.6, 6.4)],
        close=True,
        dxfattribs={"layer": "A-FLOOR-VITRIFIED"},
    )
    msp.add_text("MODULAR KITCHEN (4.8x6.2m)", dxfattribs={"layer": "ANNO", "height": 0.32}).set_placement((7.2, 3.2))

    # Master Bedroom 1 (5.2m x 5.2m)
    msp.add_lwpolyline(
        [(0.2, 6.6), (5.4, 6.6), (5.4, 11.8), (0.2, 11.8)],
        close=True,
        dxfattribs={"layer": "A-FLOOR-VITRIFIED"},
    )
    msp.add_text("MASTER BEDROOM-1 (5.2x5.2m)", dxfattribs={"layer": "ANNO", "height": 0.32}).set_placement((0.8, 9.2))

    # Bedroom 2 (4.3m x 5.2m)
    msp.add_lwpolyline(
        [(8.6, 6.6), (12.9, 6.6), (12.9, 11.8), (8.6, 11.8)],
        close=True,
        dxfattribs={"layer": "A-FLOOR-VITRIFIED"},
    )
    msp.add_text("BEDROOM-2 (4.3x5.2m)", dxfattribs={"layer": "ANNO", "height": 0.32}).set_placement((9.2, 9.2))

    # Anti-Skid Ceramic Flooring (Toilets & Utility Balcony)
    msp.add_lwpolyline(
        [(5.6, 6.6), (8.4, 6.6), (8.4, 11.8), (5.6, 11.8)],
        close=True,
        dxfattribs={"layer": "A-TOILET-FLOOR"},
    )
    msp.add_text("ATTACHED TOILETS (2.8x5.2m)", dxfattribs={"layer": "ANNO", "height": 0.28}).set_placement((5.8, 9.2))

    msp.add_lwpolyline(
        [(13.1, 6.6), (15.8, 6.6), (15.8, 11.8), (13.1, 11.8)],
        close=True,
        dxfattribs={"layer": "A-TOILET-FLOOR"},
    )
    msp.add_text("DRESS & BATH-2 (2.7x5.2m)", dxfattribs={"layer": "ANNO", "height": 0.28}).set_placement((13.3, 9.2))

    # Granite Foyer & Staircase Block
    msp.add_lwpolyline(
        [(11.6, 0.2), (15.8, 0.2), (15.8, 6.4), (11.6, 6.4)],
        close=True,
        dxfattribs={"layer": "A-GRANITE-STAIR"},
    )
    msp.add_text("FOYER & GRANITE STAIRCASE", dxfattribs={"layer": "ANNO", "height": 0.30}).set_placement((11.9, 3.2))

    # Structural RCC Columns (12 Nos)
    col_pts = [
        (0, 0), (6.5, 0), (11.5, 0), (16, 0),
        (0, 6.5), (6.5, 6.5), (11.5, 6.5), (16, 6.5),
        (0, 12), (5.5, 12), (13.0, 12), (16, 12),
    ]
    for cx, cy in col_pts:
        msp.add_blockref("RCC_COLUMN", (cx, cy), dxfattribs={"layer": "S-COLUMN"})

    # Doors (8 Nos)
    door_pts = [
        (13.5, 0), (6.5, 2.0), (11.5, 2.0), (3.0, 6.5),
        (5.5, 8.0), (8.5, 8.0), (10.5, 6.5), (13.0, 8.0),
    ]
    for dx, dy in door_pts:
        msp.add_blockref("FLUSH_DOOR", (dx, dy), dxfattribs={"layer": "A-DOOR"})

    # Windows (8 Nos)
    win_pts = [
        (3.2, 0), (9.0, 0), (0, 3.2), (0, 9.2),
        (2.8, 12), (10.8, 12), (16, 3.2), (16, 9.2),
    ]
    for wx, wy in win_pts:
        msp.add_blockref("UPVC_WINDOW", (wx, wy), dxfattribs={"layer": "A-WINDOW"})

    # Sanitary EWC Fixtures (3 Nos)
    for sx, sy in [(6.3, 10.8), (7.7, 10.8), (14.5, 10.8)]:
        msp.add_blockref("EWC_SANITARY", (sx, sy), dxfattribs={"layer": "M-SANITARY-WC"})

    doc.saveas(filepath)
    return filepath


def create_commercial_sample(filepath: str) -> str:
    doc = ezdxf.new("R2010")
    doc.header["$INSUNITS"] = 6  # Meters
    msp = doc.modelspace()

    layers = [
        ("A-WALL-EXT", 1),             # External Perimeter Masonry
        ("A-GLAZING-PARTITION", 4),    # Acoustic Glass Partitions
        ("A-OFFICE-FLOOR", 30),        # Vitrified Double-Charged Office Floor
        ("A-GRANITE-LOBBY", 6),        # Lift & Reception Granite Flooring
        ("A-FALSE-CEILING", 80),       # Gypsum + Grid Acoustic Ceiling
        ("M-WORKSTATION", 3),          # Modular Office Workstations
        ("M-CASSETTE-AC", 5),          # 2.0 TR 4-Way Cassette AC Units
        ("E-LED-PANEL", 2),            # 600x600 LED Recessed Panels
        ("A-FIRE-DOOR", 1),            # Fire & Main Doors
        ("ANNO", 7),
    ]
    for lname, col in layers:
        if lname not in doc.layers:
            doc.layers.add(lname, color=col)

    _add_block_symbol(doc, "MODULAR_DESK", radius=0.45, color=3)
    _add_block_symbol(doc, "CASSETTE_AC_2TR", radius=0.60, color=5)
    _add_block_symbol(doc, "LED_PANEL_36W", radius=0.30, color=2)
    _add_block_symbol(doc, "OFFICE_DOOR", radius=0.45, color=1)

    # Outer Core & Shell Perimeter (28m x 16m Commercial Floorplate)
    msp.add_lwpolyline(
        [(0, 0), (28, 0), (28, 16), (0, 16), (0, 0)],
        close=True,
        dxfattribs={"layer": "A-WALL-EXT"},
    )

    # Frameless Toughened Glass Cabin Partitions
    msp.add_lwpolyline(
        [(0, 11.5), (22, 11.5)],
        close=False,
        dxfattribs={"layer": "A-GLAZING-PARTITION"},
    )
    msp.add_line((7, 11.5), (7, 16), dxfattribs={"layer": "A-GLAZING-PARTITION"})
    msp.add_line((14, 11.5), (14, 16), dxfattribs={"layer": "A-GLAZING-PARTITION"})
    msp.add_line((22, 0), (22, 16), dxfattribs={"layer": "A-GLAZING-PARTITION"})

    # Open Office & Cabin Vitrified Flooring
    msp.add_lwpolyline(
        [(0.3, 0.3), (21.8, 0.3), (21.8, 11.3), (0.3, 11.3)],
        close=True,
        dxfattribs={"layer": "A-OFFICE-FLOOR"},
    )
    msp.add_text("OPEN PLAN WORKSTATION HALL (21.5x11.0m)", dxfattribs={"layer": "ANNO", "height": 0.45}).set_placement((5.5, 5.5))

    msp.add_lwpolyline(
        [(0.3, 11.6), (6.8, 11.6), (6.8, 15.7), (0.3, 15.7)],
        close=True,
        dxfattribs={"layer": "A-OFFICE-FLOOR"},
    )
    msp.add_text("BOARDROOM (16 PAX)", dxfattribs={"layer": "ANNO", "height": 0.35}).set_placement((1.2, 13.6))

    msp.add_lwpolyline(
        [(7.2, 11.6), (13.8, 11.6), (13.8, 15.7), (7.2, 15.7)],
        close=True,
        dxfattribs={"layer": "A-OFFICE-FLOOR"},
    )
    msp.add_text("MD & CEO CABIN", dxfattribs={"layer": "ANNO", "height": 0.35}).set_placement((8.5, 13.6))

    msp.add_lwpolyline(
        [(14.2, 11.6), (21.8, 11.6), (21.8, 15.7), (14.2, 15.7)],
        close=True,
        dxfattribs={"layer": "A-OFFICE-FLOOR"},
    )
    msp.add_text("SERVER & HUDDLE ZONE", dxfattribs={"layer": "ANNO", "height": 0.35}).set_placement((15.2, 13.6))

    # Granite Lift Lobby & Reception
    msp.add_lwpolyline(
        [(22.2, 0.3), (27.7, 0.3), (27.7, 15.7), (22.2, 15.7)],
        close=True,
        dxfattribs={"layer": "A-GRANITE-LOBBY"},
    )
    msp.add_text("RECEPTION & LIFT CORE LOBBY", dxfattribs={"layer": "ANNO", "height": 0.38}).set_placement((22.5, 8.0))

    # Suspended Acoustic False Ceiling across Main Office
    msp.add_lwpolyline(
        [(0.5, 0.5), (21.5, 0.5), (21.5, 15.5), (0.5, 15.5)],
        close=True,
        dxfattribs={"layer": "A-FALSE-CEILING"},
    )

    # Modular Office Workstations (24 Nos in 3 clusters of 8)
    for row_y in (2.5, 5.5, 8.5):
        for col_x in (2.5, 5.0, 7.5, 10.0, 12.5, 15.0, 17.5, 20.0):
            msp.add_blockref("MODULAR_DESK", (col_x, row_y), dxfattribs={"layer": "M-WORKSTATION"})

    # 2.0 TR 4-Way Cassette AC Indoor Units (6 Nos)
    ac_pts = [(5.0, 4.0), (11.0, 4.0), (17.0, 4.0), (3.5, 13.6), (10.5, 13.6), (18.0, 13.6)]
    for ax, ay in ac_pts:
        msp.add_blockref("CASSETTE_AC_2TR", (ax, ay), dxfattribs={"layer": "M-CASSETTE-AC"})

    # Recessed LED Panel Lights (18 Nos)
    for ly in (2.0, 6.0, 10.0):
        for lx in (3.0, 6.5, 10.0, 13.5, 17.0, 20.5):
            msp.add_blockref("LED_PANEL_36W", (lx, ly), dxfattribs={"layer": "E-LED-PANEL"})

    # Office & Fire Doors (6 Nos)
    for dx, dy in [(22.0, 3.0), (22.0, 9.0), (3.5, 11.5), (10.5, 11.5), (18.0, 11.5), (25.0, 0.0)]:
        msp.add_blockref("OFFICE_DOOR", (dx, dy), dxfattribs={"layer": "A-FIRE-DOOR"})

    doc.saveas(filepath)
    return filepath


def create_unlayered_mixed_sample(filepath: str) -> str:
    """
    Simulates a 100% LAYER-BLIND / ZERO-LAYER architectural & landscape drawing where:
    - EVERY SINGLE OBJECT is drawn on Layer '0' (zero separate layers used at all!).
    - Some rooms are drawn with closed rectangles + text (`LIVING & DINING HALL 6.60 x 11.40`),
      while other rooms (`KITCHEN 4.50 x 3.80M`, `PUJA ROOM 6'-0" x 5'-0"`, `UTILITY BALCONY 3.50 x 1.80M`)
      are written as Plan Text + Dimensions on Layer '0'!
    - Both 230mm (9") external brick walls AND 115mm (4.5") internal partition walls are drawn
      as parallel double-lines on Layer '0'!
    - Lawn, Paver Driveway, Shrub Bed, Water Pond, Trees, Palms, and Bollard Lights are ALL on Layer '0'!
    """
    doc = ezdxf.new("R2010")
    doc.header["$INSUNITS"] = 6  # Meters
    msp = doc.modelspace()

    # 1. ALL WALLS ON LAYER '0' (230mm External Double-Line Walls + 115mm Internal Partition Double-Line Walls)
    # Outer 230mm (9") double-line brick walls (gap = 0.23m) on Layer '0'
    msp.add_line((0, 0), (18, 0), dxfattribs={"layer": "0", "color": 1})
    msp.add_line((0, 0.23), (18, 0.23), dxfattribs={"layer": "0", "color": 1})

    msp.add_line((0, 12), (18, 12), dxfattribs={"layer": "0", "color": 1})
    msp.add_line((0, 11.77), (18, 11.77), dxfattribs={"layer": "0", "color": 1})

    msp.add_line((0, 0), (0, 12), dxfattribs={"layer": "0", "color": 1})
    msp.add_line((0.23, 0), (0.23, 12), dxfattribs={"layer": "0", "color": 1})

    msp.add_line((18, 0), (18, 12), dxfattribs={"layer": "0", "color": 1})
    msp.add_line((17.77, 0), (17.77, 12), dxfattribs={"layer": "0", "color": 1})

    # Internal 115mm (4.5") double-line partition walls on Layer '0' (gap = 0.115m)
    msp.add_line((7.0, 0.23), (7.0, 11.77), dxfattribs={"layer": "0", "color": 30})
    msp.add_line((7.115, 0.23), (7.115, 11.77), dxfattribs={"layer": "0", "color": 30})

    msp.add_line((7.115, 6.0), (17.77, 6.0), dxfattribs={"layer": "0", "color": 30})
    msp.add_line((7.115, 6.115), (17.77, 6.115), dxfattribs={"layer": "0", "color": 30})

    msp.add_line((13.0, 6.115), (13.0, 11.77), dxfattribs={"layer": "0", "color": 30})
    msp.add_line((13.115, 6.115), (13.115, 11.77), dxfattribs={"layer": "0", "color": 30})

    # 2. ROOMS & ZONES ON LAYER '0' (Identified purely from Plan Text & Inline Dimensions!)
    msp.add_lwpolyline(
        [(0.3, 0.3), (6.9, 0.3), (6.9, 11.7), (0.3, 11.7)],
        close=True,
        dxfattribs={"layer": "0"},
    )
    msp.add_text("LIVING & DINING HALL 6.60 x 11.40", dxfattribs={"layer": "0", "height": 0.38}).set_placement((0.8, 6.0))

    msp.add_lwpolyline(
        [(7.2, 6.2), (12.9, 6.2), (12.9, 11.7), (7.2, 11.7)],
        close=True,
        dxfattribs={"layer": "0"},
    )
    msp.add_text("MASTER BEDROOM 5.70 x 5.50", dxfattribs={"layer": "0", "height": 0.36}).set_placement((7.8, 9.0))

    msp.add_lwpolyline(
        [(13.2, 6.2), (17.7, 6.2), (17.7, 11.7), (13.2, 11.7)],
        close=True,
        dxfattribs={"layer": "0"},
    )
    msp.add_text("ATTACHED TOILET & BATH 4.50 x 5.50", dxfattribs={"layer": "0", "height": 0.34}).set_placement((13.4, 9.0))

    msp.add_lwpolyline(
        [(7.2, 0.3), (17.7, 0.3), (17.7, 5.9), (7.2, 5.9)],
        close=True,
        dxfattribs={"layer": "0"},
    )
    msp.add_text("GRANITE FOYER & STAIR LOBBY 10.50 x 5.60", dxfattribs={"layer": "0", "height": 0.36}).set_placement((8.5, 3.0))

    # 3. TEXT-ONLY DIMENSIONED ROOMS ON LAYER '0' (No closed polyline around them — recovered via Text+Dimension scanner!)
    msp.add_text("MODULAR KITCHEN 4.50 x 3.80M", dxfattribs={"layer": "0", "height": 0.36}).set_placement((3.0, 14.5))
    msp.add_text("UTILITY WET BALCONY 3.50 x 1.80M", dxfattribs={"layer": "0", "height": 0.34}).set_placement((10.0, 14.5))
    msp.add_text("STUDY ROOM 12'-0\" x 10'-0\"", dxfattribs={"layer": "0", "height": 0.35}).set_placement((16.0, 14.5))

    # 4. OUTDOOR LANDSCAPE ZONES ON LAYER '0' (Separated by Hatch Pattern & Plan Text Callouts!)
    lawn_pts = [(20, 0), (36, 0), (36, 12), (20, 12)]
    hatch_lawn = msp.add_hatch(color=3, dxfattribs={"layer": "0"})
    hatch_lawn.set_pattern_fill("GRASS", scale=0.5)
    hatch_lawn.paths.add_polyline_path(lawn_pts, is_closed=True)
    msp.add_text("FRONT LAWN TURF 16.0 x 12.0", dxfattribs={"layer": "0", "height": 0.45}).set_placement((23.5, 6.0))

    paver_pts = [(0, -5), (36, -5), (36, -1.0), (0, -1.0)]
    hatch_paver = msp.add_hatch(color=30, dxfattribs={"layer": "0"})
    hatch_paver.set_pattern_fill("HEX", scale=0.4)
    hatch_paver.paths.add_polyline_path(paver_pts, is_closed=True)
    msp.add_text("INTERLOCKING PAVER DRIVEWAY 36.0 x 4.0", dxfattribs={"layer": "0", "height": 0.45}).set_placement((11, -3.2))

    msp.add_lwpolyline(
        [(20, 13), (36, 13), (36, 15.5), (20, 15.5)],
        close=True,
        dxfattribs={"layer": "0"},
    )
    msp.add_text("PERIMETER SHRUB BED 16.0 x 2.5", dxfattribs={"layer": "0", "height": 0.4}).set_placement((23.5, 14.1))

    # 5. CIRCLES OF DIFFERENT SIZES ALL ON LAYER '0' (Auto-separated by Diameter!)
    for tx, ty in [(23, 3), (28, 3), (33, 3), (23, 9), (28, 9), (33, 9)]:
        msp.add_circle((tx, ty), radius=1.4, dxfattribs={"layer": "0", "color": 3})

    for px, py in [(19, 2), (19, 6), (19, 10), (37, 2), (37, 6), (37, 10)]:
        msp.add_circle((px, py), radius=0.9, dxfattribs={"layer": "0", "color": 80})

    for lx, ly in [(3, -0.6), (9, -0.6), (15, -0.6), (21, -0.6), (27, -0.6), (33, -0.6)]:
        msp.add_circle((lx, ly), radius=0.16, dxfattribs={"layer": "0", "color": 2})

    doc.saveas(filepath)
    return filepath


def create_sample_vector_pdf(filepath: str) -> str:
    """
    Creates a realistic 2D Vector Architectural & Landscape PDF drawing (`sample_2d_drawing_plan.pdf`)
    using PyMuPDF (`fitz`) with vector polygons, double-line walls, and text callouts so the user can
    test 2D PDF Takeoff immediately.
    """
    import fitz

    doc = fitz.open()
    # 1:100 scale where 1 meter = 28.3465 PDF points
    s = 28.3465
    page = doc.new_page(width=42 * s, height=24 * s)
    ph = 24 * s

    def pt(mx: float, my: float) -> fitz.Point:
        return fitz.Point(mx * s, ph - my * s)

    def rect(x0: float, y0: float, x1: float, y1: float) -> fitz.Rect:
        return fitz.Rect(x0 * s, ph - y1 * s, x1 * s, ph - y0 * s)

    # 1. Living & Dining Room Vector Polygon
    page.draw_rect(rect(1, 2, 8, 9), color=(0.2, 0.7, 0.9), fill=(0.1, 0.3, 0.5), width=1.5)
    page.insert_text(pt(2.0, 5.5), "LIVING & DINING HALL", fontsize=10, color=(1, 1, 1))

    # 2. Master Bedroom Vector Polygon
    page.draw_rect(rect(8.3, 2, 14.5, 9), color=(0.2, 0.7, 0.9), fill=(0.1, 0.3, 0.5), width=1.5)
    page.insert_text(pt(9.2, 5.5), "MASTER BEDROOM", fontsize=10, color=(1, 1, 1))

    # 3. Attached Toilet & Bath Vector Polygon
    page.draw_rect(rect(14.8, 2, 18.5, 9), color=(0.1, 0.5, 0.95), fill=(0.05, 0.2, 0.6), width=1.5)
    page.insert_text(pt(15.1, 5.5), "ATTACHED TOILET", fontsize=9, color=(1, 1, 1))

    # 4. Granite Foyer & Staircase Vector Polygon
    page.draw_rect(rect(1, 9.3, 18.5, 13.5), color=(0.8, 0.3, 0.8), fill=(0.4, 0.1, 0.4), width=1.5)
    page.insert_text(pt(6.0, 11.2), "GRANITE FOYER & STAIRCASE LOBBY", fontsize=10, color=(1, 1, 1))

    # 5. Front Lawn Turfing Vector Polygon (Green Fill)
    page.draw_rect(rect(20, 2, 38, 13.5), color=(0.1, 0.85, 0.3), fill=(0.1, 0.75, 0.25), width=1.5)
    page.insert_text(pt(25.0, 8.0), "CENTRAL LAWN TURFING ZONE", fontsize=11, color=(1, 1, 1))

    # 6. Interlocking Paver Walkway Vector Polygon (Orange Fill)
    page.draw_rect(rect(1, 14.2, 38, 17.2), color=(0.95, 0.55, 0.15), fill=(0.85, 0.45, 0.15), width=1.5)
    page.insert_text(pt(14.0, 15.5), "INTERLOCKING PAVER DRIVEWAY & WALKWAY", fontsize=10, color=(1, 1, 1))

    # 7. Reflective Water Body Vector Polygon
    page.draw_rect(rect(20, 18.0, 38, 22.0), color=(0.05, 0.75, 0.95), fill=(0.05, 0.45, 0.85), width=1.5)
    page.insert_text(pt(24.5, 19.8), "REFLECTIVE WATER BODY POND", fontsize=10, color=(1, 1, 1))

    # 8. Perimeter Shrub Bed Vector Polygon
    page.draw_rect(rect(1, 18.0, 18.5, 21.5), color=(0.1, 0.65, 0.4), fill=(0.08, 0.5, 0.3), width=1.5)
    page.insert_text(pt(5.5, 19.6), "ORNAMENTAL SHRUB BED", fontsize=10, color=(1, 1, 1))

    # 9. Parallel 230mm External Wall Lines
    page.draw_line(pt(0.8, 1.8), pt(18.7, 1.8), color=(0.95, 0.25, 0.25), width=1.5)
    page.draw_line(pt(0.8, 1.57), pt(18.7, 1.57), color=(0.95, 0.25, 0.25), width=1.5)

    doc.save(filepath)
    doc.close()
    return filepath


def ensure_all_samples() -> dict:
    os.makedirs(SAMPLES_DIR, exist_ok=True)
    paths = {
        "mixed": create_unlayered_mixed_sample(os.path.join(SAMPLES_DIR, "sample_mixed_unlayered_firm_style.dxf")),
        "landscape": create_landscape_sample(os.path.join(SAMPLES_DIR, "sample_landscape_masterplan.dxf")),
        "residential": create_residential_sample(os.path.join(SAMPLES_DIR, "sample_residential_villa.dxf")),
        "commercial": create_commercial_sample(os.path.join(SAMPLES_DIR, "sample_commercial_office.dxf")),
        "pdf_sample": create_sample_vector_pdf(os.path.join(SAMPLES_DIR, "sample_2d_drawing_plan.pdf")),
    }
    return paths


if __name__ == "__main__":
    created = ensure_all_samples()
    for k, v in created.items():
        print(f"Generated {k}: {v}")
