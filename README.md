# ChemDraw Energy Diagram Plotter
[![DOI](https://zenodo.org/badge/475440667.svg)](https://zenodo.org/badge/latestdoi/475440667)

This is a tool that can create pixel-accurate energy diagrams as ChemDraw objects. After opening the generated ChemDraw file, you can add chemical structures, adjust tags, and do whatever else that ChemDraw can do.

[中文版使用说明](http://bbs.keinsci.com/thread-9256-1-1.html)

## How to Use

If you're using Windows and you don't have a Python environment, you can find an executable version (packed by pyinstaller on Windows 10) in [the release](https://github.com/liyuanhe211/Energy_Diagram_Plotter_CDXML/releases/tag/3.4.1). Once you've downloaded it, go to the folder and run `Energy Diagram Plotter CDXML X.X.X.exe`.

Alternatively, you can build the virtual environment with pipenv by using the provided pipfile. And using the python environment to run Draw_Energy_Diagram_XML.py.

If you're using Linux, the script has been tested on CentOS 8 with stock Anaconda 3 2022.5.

## Introduction
To draw energy diagrams in the literature, lots of people do that by hand-drag the lines in ChemDraw and eyeball the position. But this can be extremely inaccurate and even lead to misleading results. As an example, the left figure in the image below was clearly drawn manually in ChemDraw, while the right is the to-scale version (ignore the unit change from kcal/mol to kJ/mol). The original diagram is a mess and even worse than just giving numbers as a table. There are other tools available that can generate energy diagrams, such as DataGraph, mechaSVG, and Origin, but they usually generate an un-editable figure with very limited customization options. Additionally, you cannot drag the numbers and tags, which can be troublesome for complex energy diagrams.

<p align="center"><img src="https://user-images.githubusercontent.com/18537705/160632947-6754c8b0-a5f2-45d3-9577-d1a10f9f4ea8.png" width="100%" height="100%" align="center"></img></p>

The program here allows you to genreate pixel-accurate energy diagrams as ChemDraw objects, Below is the input file format, and there are also [several several other examples](https://github.com/liyuanhe211/Energy_Diagram_Plotter_CDXML/tree/main/Examples) available in this repository:

<p align="center"><img src="https://user-images.githubusercontent.com/18537705/160620003-5657e605-e95c-495b-aeae-b43006e78b6b.png" width="70%" height="70%" align="center"></img></p>

<p align="center"><img src="https://user-images.githubusercontent.com/18537705/160621422-05274905-5b1e-43b7-8cde-1ae80577d795.png" width="100%" height="100%" align="center"></img></p>

You can use any color for each state with its RGB value #XXXXXX. Some colors have shortcuts, and the "tabcolors" can be called by **B**, **O**, **G**, **R**, **P**, **BR**, **PI**, **GR**, **OL**, **C** (all uppercase):

<p align="center"><img src="https://user-images.githubusercontent.com/18537705/175197188-bee314cd-73fb-4918-81ac-9fa89b47cc9f.png" width="30%" height="30%" align="center"></img></p>

<p align="center"><img src="https://user-images.githubusercontent.com/18537705/185797574-d620194d-6558-423e-ba4f-a1d7c04b63b2.png" width="60%" height="60%" align="center"></img></p>

There are also some predefined "pure colors" that can be used with the shorthand codes **y**, **m**, **c**, **r**, **g**, and **b** (top to bottom).

<p align="center"><img src="https://user-images.githubusercontent.com/18537705/160639914-d11c34dc-5c1d-486e-b6b4-f8216883adba.png" width="70%" height="70%" align="center"></img></p>

The tool also provides several options to customize the energy diagram. The options in the GUI are explained in the following figures (You can also hover on the options to see a tooltip):

<p align="center"><img src="https://user-images.githubusercontent.com/18537705/160637389-257753fc-e0bc-4822-9d44-c86e0c8207a0.png" width="100%" height="100%" align="center"></img></p>

(The font of the figure is fixed to Arial (you can change it later in ChemDraw), and the canvas size and aspect ratio of the figure are determined by the window size.)

Here are some visuals for each option. Note that some of the following options are only reflected in the ChemDraw file, not in the preview view:

<p align="center"><kbd><img src="https://user-images.githubusercontent.com/18537705/160631167-2ffa2ade-e520-4bd6-86f2-c637a86716f4.png" width="70%" height="70%" align="center"></img></kbd></p>

<p align="center"><kbd><img src="https://user-images.githubusercontent.com/18537705/160631129-c5988851-7960-4817-abbc-8190882fb4b4.png" width="70%" height="70%" align="center"></img></kbd></p>

<p align="center"><kbd><img src="https://user-images.githubusercontent.com/18537705/160631058-3d0f9959-e18f-48c2-b42d-a3cd2005f035.png" width="70%" height="70%" align="center"></img></kbd></p>

<p align="center"><kbd><img src="https://user-images.githubusercontent.com/18537705/160630699-5f7ff0a1-a0d6-4e36-b200-fd9bbcc42601.png" width="70%" height="70%" align="center"></img></kbd></p>

<p align="center"><kbd><img src="https://user-images.githubusercontent.com/18537705/160630414-0c20936a-b9e7-41be-ad61-c7a1b3036394.png" width="70%" height="70%" align="center"></img></kbd></p>

The last option, "Use Temporary File," allows you to choose whether the program creates a file in its own Temp folder (which you need to Save As to another place after viewing) or directly generates a CDX file in the directory where Excel is located.

Other options that's not mentioned above should be self-explanatory.

## Known issues
 * The program stores its temporary files in the program directory, so it requires read and write permissions to this directory. It's not recommended to install the program in a path with permission restrictions, such as Program Files. If you do so, please grant the program administrator permissions.

 * The preview interface is not entirely accurate. If you experience any unexpected behavior, export the ChemDraw file first to check if it's normal.

 * Changes to the format (title, label, ticks, etc.) through the Matplotlib toolbar will not be carried over to the exported CDXML file (except for adjustments to the plot region). Please make those adjustments in ChemDraw.

 * The "Avoid Text Overlap" function uses a greedy algorithm and may produce unreasonable results in particularly crowded situations. You can fine-tune it yourself by holding down Shift and dragging the text in ChemDraw to ensure it moves horizontally or vertically.

<p align="center"><kbd><img src="https://user-images.githubusercontent.com/18537705/160629770-3f10e450-e240-4ac5-96f0-0e80434bd413.png" width="70%" height="70%" align="center"></img></kbd></p>


## For bug report

If you encounter any unexpected behavior, please try running the [examples](https://github.com/liyuanhe211/Energy_Diagram_Plotter_CDXML/tree/main/Examples) first to determine whether the problem is with your input or the program.

If the issue is with the program, please provide feedback by uploading your Excel input file, a screenshot of the GUI before the crash, and the last display in the CMD window. If the output file is generated but there is a problem with it, please upload the output file as well.

If the program crashes, please run it using the "Draw_Energy_Diagram_XML.bat" file in the directory to keep the error message before the program exits.

## Citation

Citing this program is optional. If you choose to cite this, you can use the Zenodo DOI: `10.5281/zenodo.7187658`. 

E.g. Li, Y.-H. _Energy Diagram Plotter CDXML 3.5.1_ (DOI: 10.5281/zenodo.7187658), **2022**.
