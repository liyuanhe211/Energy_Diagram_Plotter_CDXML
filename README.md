# Energy Diagram Plotter CDXML
A tool to create pixel-accurate energy diagrams as ChemDraw object.

If you don't have Python environment, there is an executable version (packed by cx_freeze on Windows 10) in [the release](https://github.com/liyuanhe211/Energy_Diagram_Plotter_CDXML/releases/tag/3.4).

There is also a [Chinese version](http://bbs.keinsci.com/thread-9256-1-1.html) of this article.

## Background
In the literatures, I see lots of people use the ChemDraw+mouse+eye method to draw energy diagrams. It is tedious and sometimes extremely inaccurate. Sometimes even the qualitative order breaks down, which could be very misleading. As an example, the left figure below is from a literature which is clearly drawn manually in ChemDraw, and the right is the to-scale version (ignore the unit change from kcal/mol to kJ/mol). The original diagram is a mess. Such diagrams are even worse than only giving numbers as a table.

<p align="center"><img src="https://user-images.githubusercontent.com/18537705/160632947-6754c8b0-a5f2-45d3-9577-d1a10f9f4ea8.png" width="100%" height="100%" align="center"></img></p>

There are lots of tools to generate energy diagrams. For example, DataGraph or [mechaSVG](https://github.com/ricalmang/mechaSVG). They are OK for basic graphs, but they usually generate an un-editable figure with very limited customization options. Also, one cannot drag the numbers and tags, which is troublesome for complex energy diagrams.

## This program
The program here can automatically generate ChemDraw files with line charts based on simple Excel input. After that, you can freely modify it in ChemDraw. And it's very convenient to add chemical structure:

<p align="center"><img src="https://user-images.githubusercontent.com/18537705/160620003-5657e605-e95c-495b-aeae-b43006e78b6b.png" width="70%" height="70%" align="center"></img></p>

Below is the input file format. By comparing left to right, it should be easy to understand. I also give [several other examples](https://github.com/liyuanhe211/Energy_Diagram_Plotter_CDXML/tree/main/Examples) in this repo:

<p align="center"><img src="https://user-images.githubusercontent.com/18537705/160621422-05274905-5b1e-43b7-8cde-1ae80577d795.png" width="100%" height="100%" align="center"></img></p>

For color definition, you can use any color for each state with its RGB value `#XXXXXX`. The colors below it can also be used with `y`,`m`,`c`,`r`,`g`,`b` (top to bottom).:

<p align="center"><img src="https://user-images.githubusercontent.com/18537705/160639914-d11c34dc-5c1d-486e-b6b4-f8216883adba.png" width="70%" height="70%" align="center"></img></p>

The overall GUI structure and the meaning of each option are described in the following figures (You can also hover on the options to see a tooltip):

<p align="center"><img src="https://user-images.githubusercontent.com/18537705/160637389-257753fc-e0bc-4822-9d44-c86e0c8207a0.png" width="100%" height="100%" align="center"></img></p>

(The font of the figure is fixed to Arial (you can change it later in ChemDraw), and the canvas size and aspect ratio of the figure are determined by the window size.)

Here are some visuals for each option. Note that some of the following options are only reflected in the ChemDraw file, not in the preview view:

<p align="center"><kbd><img src="https://user-images.githubusercontent.com/18537705/160631167-2ffa2ade-e520-4bd6-86f2-c637a86716f4.png" width="70%" height="70%" align="center"></img></kbd></p>

<p align="center"><kbd><img src="https://user-images.githubusercontent.com/18537705/160631129-c5988851-7960-4817-abbc-8190882fb4b4.png" width="70%" height="70%" align="center"></img></kbd></p>

<p align="center"><kbd><img src="https://user-images.githubusercontent.com/18537705/160631058-3d0f9959-e18f-48c2-b42d-a3cd2005f035.png" width="70%" height="70%" align="center"></img></kbd></p>

<p align="center"><kbd><img src="https://user-images.githubusercontent.com/18537705/160630699-5f7ff0a1-a0d6-4e36-b200-fd9bbcc42601.png" width="70%" height="70%" align="center"></img></kbd></p>

<p align="center"><kbd><img src="https://user-images.githubusercontent.com/18537705/160630414-0c20936a-b9e7-41be-ad61-c7a1b3036394.png" width="70%" height="70%" align="center"></img></kbd></p>


The last "Use temp file" option is to decide whether the program creates a file in its own Temp folder (the user needs to save it after viewing it), or directly generates a cdx file in the directory where Excel is located.

Other options not mentioned should be self-explanatory.

## Other issues

The temporary files of the program are stored in the program directory, so the program must have read and write permissions to this directory. It is not recommended to put the program in a path with permission restrictions such as Program Files, otherwise please give it administrator permissions.

The preview interface is not completely correct. If there is unexpected behavior, export the ChemDraw file first to see if it is normal.

The "Avoid text overlap" function is currently a greedy algorithm, and may produce unreasonable results in particularly crowded situations. Just fine-tune it yourself (tip, hold down Shift and drag the text in ChemDraw to ensure it moves horizontal or vertical).

<p align="center"><kbd><img src="https://user-images.githubusercontent.com/18537705/160629770-3f10e450-e240-4ac5-96f0-0e80434bd413.png" width="70%" height="70%" align="center"></img></kbd></p>


## For bug report

If you encounter some unexpected behavior, please try to run the [examples](https://github.com/liyuanhe211/Energy_Diagram_Plotter_CDXML/tree/main/Examples) first and see whether it's a problem of your input or my program.

After that, you can give feedback by upload your Excel input file, a screenshot of the operation interface before crashing and the last display in the CMD window. If the output file is generated, but there is a problem with the output file, please upload the output file as well. 

If there is a crash, please run the program through the "Draw_Energy_Diagram_XML_Debug.bat" file in the directory, to keep the error message before the program exits. 
