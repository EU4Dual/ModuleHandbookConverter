#!/bin/bash

# Set the source and destination directories
source_dir="pdf"
dest_dir="html"

# Check if the destination directory exists, if not, create it
sudo mkdir -p "$dest_dir"

# Loop through all PDF files in the source directory
for pdf_file in "$source_dir"/*.pdf; do
    # Get the base filename without extension
    base_filename=$(basename -- "$pdf_file")
    base_filename_noext="${base_filename%.*}"

    # Run pdf2htmlEX on each PDF file
    pdf2htmlEX --embed cfijo --dest-dir "$dest_dir" "$pdf_file"

    # Optionally, you can print a message for each conversion
    echo "Converted $pdf_file to HTML: $dest_dir/$(basename "$pdf_file" .pdf).html"
done

echo "Batch conversion completed."
