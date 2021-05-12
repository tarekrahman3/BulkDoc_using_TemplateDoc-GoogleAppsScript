#!/bin/bash

for file in ./*.docx; do
    [ -e "$file" ] || continue
    lowriter --headless --convert-to pdf  "$file"
done


