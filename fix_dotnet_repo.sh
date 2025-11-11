#!/bin/bash
# Fix .NET repository issue for Ubuntu Focal
# Run with: bash fix_dotnet_repo.sh

echo "Fixing .NET repository configuration..."

# Remove the broken backports repository
sudo rm -f /etc/apt/sources.list.d/dotnet-ubuntu-backports-focal.list

# Add Microsoft's official repository
wget https://packages.microsoft.com/config/ubuntu/20.04/packages-microsoft-prod.deb -O /tmp/packages-microsoft-prod.deb
sudo dpkg -i /tmp/packages-microsoft-prod.deb
rm /tmp/packages-microsoft-prod.deb

# Update package list
sudo apt-get update

# Install .NET SDK 9.0 (or specify a different version)
echo "To install .NET SDK 9.0, run:"
echo "  sudo apt-get install -y dotnet-sdk-9.0"
echo ""
echo "Or for .NET SDK 8.0 (LTS):"
echo "  sudo apt-get install -y dotnet-sdk-8.0"
echo ""
echo "Or for .NET SDK 6.0 (LTS):"
echo "  sudo apt-get install -y dotnet-sdk-6.0"

echo ""
echo "Repository fix complete! You can now install the .NET SDK."


