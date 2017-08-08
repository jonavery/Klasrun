# Overview
Every script in the Klasrun library serves the purpose of saving time and improving accuracy. The scripts accomplish this through automating tasks that would otherwise be menial, arduous, and time-consuming. There are essentially two halves to the Klasrun automation library - Google Sheets and Amazon Marketplace Web Service (MWS). Google Sheets serves as our online inventory dataplace, while we use MWS to take that inventory and sell it to customers across North America. The code in the Klasrun library makes it possible for these two halves to directly communicate with each other.

# Data Flow
All inventory data originates from a Blackwrap Manifest.  The manifest is imported into the Liq Orders sheet. We research the items and designate them as Amazon, Ebay, or Return, then send them to the Manifest Load sheet. This sheet formats the inventory to get it ready to go into our database, then sends it to both the Liquidation and Work sheets. The inventory is highlighted by designation in the Future tab, then sent to the Listings tab. It is tested and rated, then sent to the Scrap tab. The information is audited for completeness and accuracy, then sent to the MWS and New Archive tabs. The MWS tab then prices out the items, lists them on Amazon, and puts them into shipments.

# Manifest
The manifest section of the database is entirely concerned with inbound auctions and blackwraps. It handles everything from the moment we get a blackwrap manifest up to the inventory physically moving into the warehouse and digitally into the work and liquidation sheets.

Importing Blackwrap Manifest
1. Open new blackwrap manifest
2. Copy sheet into ‘liq orders1.ods’
3. In liq orders1 sheet Automation Menu: Import Blackwrap, Generate Prices, Import Prices
4. Name blackwrap in ‘Sheet6’ and fill out information in ‘AUCTION’

# Inventory Processing
All inventory processing occurs within the various tabs of the WORK sheet. First, new inventory arrives in the Future Listings tab. There, it is highlighted according to its pre-designation. After that, it is moved into the Listings tab. Inventory is tested, graded, and evaluated. Completed inventory is moved either to the EBAY sheet, SCRAP tab, or New Archived tab depending on its designation. Inventory destined for shipment to Amazon is packed, weighed, and measured. Completed packages are copied into the MWS tab, where they are priced and sent to Amazon. The packages are then archived from the Scrap tab into the New Archived tab.

Import Inventory to Listings
1. Open WORK
2. Navigate to Future Listing
3. Run ‘Highlight Future Listings by A/E/R’ in Automation Menu
4. Cut & paste inventory into Listings

Add Extra Items
1. Run ‘Get Next SKU From Liquidation’ in Automation Menu to get next SKU
2. Type in new SKU, Item Title, ASIN, A/E/R, & ORDER # on next line of Listings
3. Run either ‘Update Item in Liquidation by SKU’ or ‘Update All Work Items in Liquidation’

# Listing Inventory on Amazon
For items to be sold on Amazon, they must first be listed. Klasrun has scripts that enable listing to be done straight from the Google Sheets inventory database. The Klasrun administrator moves inventory from the Listings tab to the SCRAP tab after testing is complete. Once inventory has been packaged and dimensions are entered, it is ready to be listed on Amazon. The script is set-up so that only items with non-blank weights will be listed. To list electronic or palleted items (which do not have dimensions), put an ‘x’ in the item’s weight column. It is usually best practice to list each type of item and then create shipments for it before listing the next type.

# Future Improvements
- [ ] Create names (e.g. ‘BLACKWRAP 9’) in liq order Sheet6 during the ‘Import New Blackwrap’ script.
- [ ] Remove MANIFEST LOAD sheet and move its scripts into the liq order sheet. Probably rename liq orders1.ods to Manifest Load in this scenario.
- [ ] When generating shipments, add ability for user to select Electronics (all items in one box), Pallet (LTL instead of SmallParcel), or Regular.
- [ ] Audit WORK Listings sheet. Highlight incomplete inventory cyan and move completed inventory to the SCRAP and New Archived tabs as appropriate.
