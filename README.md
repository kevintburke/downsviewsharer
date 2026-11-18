# downsviewsharer
Tool for processing titles destined for Keep@Downsview SHARE program. Includes checks for valid data, items on loan/in process, and steps for updating bibs, holdings, and item records.

## Background
This tool accepts spreadsheets of items intended for uOttawa's K@D SHARE program and performs the necessary metadata updates for that workflow. Spreadsheets must be in .xlsx or .csv format and include at least columns for uOttawa's MMS ID, UofT's MMS ID, and uOttawa's Item PID. Additionally, users will need an API Key for Alma API in the production or sandbox environment of their choice.

## Inputs
This tool requires two inputs
1. A spreadsheet (.xlsx or .csv) with columns for uOttawa MMS Id, uOttawa Item PID, and UofT MMS ID
2. An API Key for Alma API

## Special Notes
This script is currently set up for uOttawa processes only, meaning that barcode checks, location names and codes, and local field updates are specific to uOttawa's workflows. The script will need substantial revision for use at other institutions.
<del>Part of this script relies on PyZ3950 to retrieve records from UofT's Z39.50 server. PyZ3950 appears to no longer be actively maintained and there are some code breaking issues that require correction before it can be used. There are notes about the corrections required in comments at the top of the script, but these are not comprehensive. Note that the changes recommended may break certain functionality for other projects, but enable sufficient functioning for this workflow.</del><ins>This portion of the script has been disabled as the data retrieved from UofT's Z39.50 server are not required for validating SHAREs and may be overwritten.</ins>

## Process
The script follows this general outline:
1. Checks the current item records in Alma for Annex-type barcodes, items on-loan, or items otherwise in a process; removes them from processing, and adds them to lists for manual review at the end
2. Checks UofT's bib record for 949 $l REMOTEOS and removes item from processing if not found, adding to a list for manual review
3. Updates our item record to...
    - Strip the barcode
    - Change the location to xdownsview
    - Add a retention commitment, note, and reason
4. Retrieves the holdings record associated with the updated item (this may have changed if we're splitting holdings across locations) and adds an 014 with the UofT MMS ID if it doesn't already exist
5. Updates our bib record to add a 940 field, if one doesn't exist

Each stage of the process that requires API or Z39.50 access implements an exponential backoff of up to 64 seconds before abandoning the attempt. Error handling is minimal, and any failed attempts to retrieve or update records will have those records' identifiers (MMS, Holdings, or Item PID, depending on the type of record) added to a list of errors saved on completion.
