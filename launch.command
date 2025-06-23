#!/bin/bash

# Activate conda environment
source /Users/cbrichet/anaconda3/etc/profile.d/conda.sh
conda activate streamlitenv

# Get the directory of the script
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

# Change to the script directory
cd "$DIR"

# Run the app with streamlit
streamlit run app.py 