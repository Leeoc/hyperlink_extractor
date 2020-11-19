#!/usr/bin/env python
import os
import PySimpleGUI as sg
import pandas as pd
import numpy as np

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT


def main():
    sg.set_options(auto_size_buttons=True)

    filename = sg.popup_get_file(
        'filename to open', no_window=True, file_types=(("Word Documents", "*.docx"),))
    if filename == '':
        return

    data = []
    header_list = []

        # --- populate table with file contents --- #
    if filename is not None:
        try:
            # Load the data
            df, col_len, link_count = get_hyperlinks(filename)
            data = df.values.tolist() 

            # Create headers for the table
            if data:
                header_list = ['Link', 'Count']
            else: 
                # Creates columns names for each column ('column0', 'column1', etc)
                header_list = ['column' + str(x) for x in range(len(data[0]))]
        except:
            sg.popup_error('No links found.')
            return

    layout = [
        [sg.Table(values=data,
                  headings=header_list,
                  display_row_numbers=False,
                  col_widths=[150, 3],
                  auto_size_columns=True,
                  enable_events=True,
                  num_rows=min(25, len(data)))],
                  [sg.Text("{} links found".format(link_count))]
    ]

    # Display the window and check for events
    window = sg.Window('HyperLink Extractor',layout, grab_anywhere=False)
    while True:
        event, values = window.read()
        # On click copy the link to clipboard 
        for i in range(len(data)):
            if {0: [i]} == values:
                print(data[i][0])
                addToClipBoard(data[i][0])
        if event == "OK" or event == sg.WIN_CLOSED:
            break
    window.close()

def get_hyperlinks(path):
    document = Document(path)

    rels = document.part.rels
    link_count = 0 
    link_list = []
        
    for rel in rels:
        if rels[rel].reltype == RT.HYPERLINK:
            link_count += 1
            link_list.append([rels[rel]._target, link_count])

    df = pd.DataFrame(data=link_list, columns=['HyperLink' , 'Count'])
    print(df.head(100))

    return df, longestString(df), link_count

def longestString(df):
    max_len = 0
    for row in df['HyperLink']:
        row_len = len(row)
        if row_len >= max_len:
            max_len = row_len
    return max_len

def addToClipBoard(text):
    command = 'echo ' + text.strip() + '| clip'
    os.system(command)
    return


if __name__ == '__main__':
    main()

