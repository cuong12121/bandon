"""
Configuration for barcode define characters and this machine's define.

Edit `CURRENT_DEFINE` to match the character for this machine (e.g. '$' or '&').
VALID_DEFINES lists the allowed define characters.
"""

# Allowed define characters that may appear at the end of a scanned barcode
VALID_DEFINES = ['#', '$', '%', '&', '*']

# The define character that this machine/listener should react to.
# Change this value per-machine to control which barcodes cause a cut.
CURRENT_DEFINE = '$'
