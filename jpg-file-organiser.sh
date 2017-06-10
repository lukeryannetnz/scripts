#!/bin/bash

# Organises all JPG files into folders based on when their file attributes were last set
echo Welcome to the file organiser. I will organise all jpg files into folders based on their ctime Year and Month.

echo -n "Enter the root directory name to organise files into: "
read rootDir
mkdir $rootDir

while IFS= read -r -d $'\0' file; do
    IFS=$'\n'

    echo $file

    date=$(stat -f "%Sm" -t "%Y-%m" $file)
    currentDir=$rootDir/$date
    mkdir -p $currentDir
    cp $file $currentDir/$(basename $file)

    #create an output file with the original filenames
    echo $file >> $currentDir/originalFileNames.txt

#find files that aren't in the rootDir which have the jpg extension. Redirect them to the while loop.
done < <(find . -path ./$rootDir -prune -or -iname "*.JPG" -type f -print0)

unset IFS
