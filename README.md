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
