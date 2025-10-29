# VisualizeAutomate
The Standard version of Visualize does not have access to the API, so created this Python script that records the step by step process of the rendering process. 
## Command Format
python visualize_automator.py --exec "<path_to_your_VAULT_export.xlsx>" --vault "vault_name"
## Where the script expects to find the CAD models:
C:\AVPVault\Genie\Design Engineering\Library\

# 🟦 Vault Mode Column Requirements

The sheet you pass in must contain:

## Column	Meaning	Example
A	Part Number (Item)	19136261
B	Short Description	Support Bracket
C (if present)	Revision	A, B, etc.

### The script automatically does:

<Item>.SLDPRT
<Item>.SLDASM
<Item>.SLDDRW
<Item>.SLDPRT / with GT, SGT, PLT suffix cleanup

## 🟦 Vault Mode Quick Debug Commands
Validate Vault file lookup without opening Visualize:
python visualize_automator.py --vault-test "<excel_file_path>"

### Just show the matched Vault file paths:
python visualize_automator.py --vault-print "<excel_file_path>"

## 🟦 The Other Mode for Comparison (CSV Mode)

If you already have absolute paths and don't want vault lookup:

python visualize_automator.py --exec "<some_regular_image_or_cad_list.csv>"

Mode	          Use When
--vault	        You exported from PDM Search and want it to auto-locate files
--exec only	    You already have direct folder paths in the CSV
