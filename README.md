# mortpak-script

This is a python program for parsing Mortpak's output report into a more versatile csv format.

## Requirements

- Mortpak for Windows

- Foxit Reader

- Mircrosoft Excel 2016

## How to

These instructions work with "Mortpak for Windows", "Foxit Reader", Microsoft Excel 2016

Download the contents of this repository

Unzip the files somewhere

The folder will be called `mortpak-script-master`

Inside there is a folder called `mortpak`. *This contains the main program and it is where we will save our Excel Workbook later.*

*In Mortpak*

`File` > `Open` > (Open your `.mpl` file)

`Run` application

`File` > `Print Document Output`

Choose `Microsoft Print to PDF` > `Print`

`Save` it under any name

*In Foxit Reader*

Open the PDF in Foxit Reader

Select all (`Ctrl + A`)

Copy (`Ctrl + C`)

*In Excel 2016*

Open a new Excel workbook

Use `Ctrl + V` to paste into cell `A1`

Save the workbook in the `mortpak` folder within the `mortpak-script-master` folder from earlier

Save the workbook under any name, but make sure it is saved as `Excel Workbook` type

Open the folder named `mortpak`

*In mortpak folder*

Go to the `mortpak` folder (where we just saved the Excel Workbook)

Open the `mortpak.exe` application (your computer may just display it as `mortpak`)

When prompted, enter the name of the Excel workbook (no need for `.xlsx` suffix, just the name will do) and press `Enter`

In `mortpak` folder, you will find `mortpak_results.csv` file

Fini!
