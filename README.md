# AutoBOQ 2D — On-Screen 2D Takeoff, BOQ, Abstract & Bar Bending Schedule (IS 1200 / IS 2502)

A fast, layer-independent 2D Drawing Quantity Takeoff, Bill of Quantities (BOQ), Abstract of Cost, and Bar Bending Schedule (BBS) web application for **Landscape, Residential, and Commercial** projects.

## Key Features
- **True 2D Drawing Viewer (`.DXF`, `.DWG`, `.PDF`)**: Renders drawings faithfully using `ezdxf` & `PyMuPDF`—works even when all items are on a single layer or unlayered.
- **On-Screen Precision Takeoff**: Length, Area Polygon, Rectangle (`L × B`), Count stamps, Text-dimension picker (`4.50 x 3.60`), Vertex Snapping, Ortho, and 2-point Scale Calibration.
- **IS 1200 / CPWD DSR BOQ & Abstract**: Live Measurement Sheet, Bill of Quantities, and Cost Abstract with Contingency & GST.
- **IS 2502 / SP 34 Bar Bending Schedule (Footing, Column, Beam, Slab)**:
  - Footing **Bottom Mesh (`Lx` & `Ly`)** and **Top Mat (`Lx` & `Ly`)**
  - Dynamic **Shape Codes (`SC-00`, `SC-01`, `SC-02`, `SC-41`, `SC-51`)** & **Bend / Crank Angles (`0°`, `45°`, `90°`, `135°`, `180°`)**
  - Visual detailing guides for **Cranked Bar Arrangement**, **Closed Stirrup $135^\circ$ Hook Detail**, and **Extra Support Bars**
  - Standalone 6-sheet Excel BBS workbook (`Bar_Bending_Schedule_Footing_Column_Beam_Slab.xlsx`) with native Linked Pictures.
- **1-Click Excel (`.xlsx`) & PDF (`.pdf`) Export**: Uses Microsoft Excel on Windows or LibreOffice Calc (`soffice`) on Linux cloud hosts.

## Run Locally
```bash
pip install -r requirements.txt
uvicorn app:app --host 127.0.0.1 --port 8095
```
Open http://127.0.0.1:8095 in your browser.

## Deploy to Cloud (Render / Railway / Google Cloud Run)
A `Dockerfile`, `Procfile`, and `requirements.txt` are included. Any Docker-capable cloud platform will automatically install Python + LibreOffice Calc for headless PDF export and expose the app on `$PORT`.
