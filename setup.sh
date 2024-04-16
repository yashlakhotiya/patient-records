#!/bin/bash

# Create and activate virtual environment
# chmod +x setup.sh
# ./setup.sh
python3 -m venv venv
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

source venv/bin/activate
echo "Setup complete."
