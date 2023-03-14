# Document-Search-Program.py
Document Search Program

This code creates a graphical user interface (GUI) application that allows the user to search for specific keywords across all .docx files in a selected folder, and highlights the keywords in the matching files.

Here's a breakdown of what the code does:

It imports several libraries, including os, tkinter, subprocess, docx2txt, docx, keyword, string, glob, PIL, and enum.

It defines an enumeration class Color to define color values, and a function highlight_keyword() to highlight keywords in a given .docx file.

It defines a function extract_text_from_docx() that extracts text from a given .docx file.

It creates a tkinter window and several widgets, including buttons, labels, and listboxes.

It binds various functions to different events, such as button clicks or key presses.

It defines a function search_folder() that searches for specific keywords in all .docx files in a selected folder and displays the matching files in a tkinter listbox.

It defines a function open_file() that opens the selected matching file and highlights the searched keywords.

It defines a function reset_data() that resets the selected folder and keyword fields.

It sets up the layout and structure of the GUI application.

Finally, it starts the main event loop to run the GUI application.

Overall, this code creates a simple but useful GUI application that can help users search for specific keywords in a large collection of .docx files, and quickly navigate to the relevant files.
