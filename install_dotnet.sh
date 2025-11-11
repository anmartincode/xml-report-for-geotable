#!/bin/bash
# Install .NET SDK 8.0 on Ubuntu 20.04

echo "Installing .NET SDK..."

# Add Microsoft package repository
wget https://packages.microsoft.com/config/ubuntu/20.04/packages-microsoft-prod.deb -O packages-microsoft-prod.deb
sudo dpkg -i packages-microsoft-prod.deb
rm packages-microsoft-prod.deb

# Update package list
sudo apt-get update

# Install .NET SDK
sudo apt-get install -y dotnet-sdk-8.0

# Verify installation
echo ""
echo "Verifying installation..."
dotnet --version

echo ""
echo "✓ .NET SDK installed successfully!"
echo ""
echo "You can now build the project with: dotnet build"
