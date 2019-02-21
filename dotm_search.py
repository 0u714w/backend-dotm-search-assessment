#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
import os
import zipfile
import argparse
import glob

__author__ = "dougenas"


def search_text(character, path):
    """Searches a directory for .dotm files, then searches that file for a specific character"""
    print('Searching directory {} for text {}').format(path, character)
    files_searched = 0
    files_matched = 0
    os.chdir(path)
    files = glob.glob('*.dotm')
    for file in files:
        files_searched += 1
        zipped_file = zipfile.ZipFile(file, 'r')
        zip_read = zipped_file.read('word/document.xml')
        if character in zip_read:
            index = zip_read.index(character)
            files_matched += 1
            output = 'Match found in file ' + file + '\n' + '...{}...'.format(zip_read[index - 40: index + 40])
            print(output)

    print('Total files searched: {}').format(files_searched)
    print('Total items matched: {}').format(files_matched)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--dir', help='add a directory')
    parser.add_argument('character', help='search for a text item')
    args = parser.parse_args()
    character = args.character
    dir = args.dir
    if dir:
        search_text(character, dir)
    else:
        print('{} is not a directory').format(args.dir)
        print('Please provide a directory to be searched.')


if __name__ == '__main__':
    main()
