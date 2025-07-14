# Slide-Inventory-Script

Simple Python utilities that turn a folder (and its sub‑folders) of  
`.tif/.tiff` files into a clean **Excel workbook** ready for cataloguing.

Key features
------------

* **Natural sort** – filenames like `…-01_15`, `…-02_01`, `…-10_01` appear in true numeric order.  
* **Drop‑down lists** – validated “Format” and “Extent” columns to speed data entry.  
* **Optional auto‑classification** – supply regex ➜ (Format, Extent) rules and the script pre‑fills those columns.  
* Pure Python 3; only external dependency is **`openpyxl`**.

### Project background

This utility was created on my own initiative while working as a **Digital
Collections Student Assistant** at the CSUN Library.  The library’s workflow
did not include an automated inventory step, so I built this script to
eliminate the manual task of copying hundreds of TIFF filenames into Excel and
adding metadata drop‑downs.  It is now shared here as an open‑source tool that
other archives can adapt.
