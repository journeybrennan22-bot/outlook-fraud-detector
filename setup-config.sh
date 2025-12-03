#!/bin/bash
# setup-config.sh - Configure the Outlook add-in with your values

echo "=== Email Fraud Detector - Configuration Setup ==="
echo ""

# Get user input
read -p "Enter your hosting domain (e.g., yourusername.github.io/repo): " DOMAIN
read -p "Enter your Azure AD App Client ID: " CLIENT_ID

# Validate inputs
if [ -z "$DOMAIN" ] || [ -z "$CLIENT_ID" ]; then
    echo "Error: Both domain and client ID are required"
    exit 1
fi

# Generate new GUID for the add-in
NEW_GUID=$(cat /proc/sys/kernel/random/uuid 2>/dev/null || uuidgen 2>/dev/null || echo "GENERATE-NEW-GUID-MANUALLY")

echo ""
echo "Updating configuration files..."

# Update manifest.xml
sed -i "s|YOUR_DOMAIN|$DOMAIN|g" manifest.xml
sed -i "s|YOUR_APP_CLIENT_ID|$CLIENT_ID|g" manifest.xml
sed -i "s|a1b2c3d4-e5f6-7890-abcd-ef1234567890|$NEW_GUID|g" manifest.xml

# Update taskpane.js
sed -i "s|YOUR_APP_CLIENT_ID|$CLIENT_ID|g" src/taskpane.js
sed -i "s|YOUR_DOMAIN|$DOMAIN|g" src/taskpane.js

echo ""
echo "=== Configuration Complete ==="
echo ""
echo "Domain:    https://$DOMAIN"
echo "Client ID: $CLIENT_ID"
echo "Add-in ID: $NEW_GUID"
echo ""
echo "Next steps:"
echo "1. Add your icon files to the assets/ folder"
echo "2. Host the files at https://$DOMAIN"
echo "3. Sideload or deploy via Microsoft 365 Admin Center"
echo ""
