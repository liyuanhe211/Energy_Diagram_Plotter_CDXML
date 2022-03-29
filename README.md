# Energy Diagram Plotter CDXML
A tool to create pixel-accurate energy diagrams as ChemDraw object.

## Background
In the literatures, I see lots of people use the ChemDraw+mouse+eye method to draw energy diagrams. It is tedious and sometimes extremely inaccurate. Even the order of energy could be accidentally inverted. Which could be very misleading. The left figure below is the line graph drawn by ChemDraw in the literature, and the right is the to-scale version (ignore the unit change from kcal/mol to kJ/mol). The original diagram is a mess. Such diagrams are even worse than only giving numbers as a table.

![image](https://user-images.githubusercontent.com/18537705/160621653-d8dd5409-9351-4fda-81aa-b602a1f8afd2.png)

There are lots tools to generate energy diagrams. For example DataGraph or [mechaSVG](https://github.com/ricalmang/mechaSVG). They are OK for basic graphs, but they usually genreates an un-editable figure with very limited customization options. Also one cannot drag the numbers and tags, trubblesome for complex energy diagrams.

#
## This program
The program here can automatically generate ChemDraw files with line charts based on simple Excel input. After that, you can freely modify the ChemDraw elements. And it's very convient to add chemical structure to it.

![image](https://user-images.githubusercontent.com/18537705/160620003-5657e605-e95c-495b-aeae-b43006e78b6b.png)

The input file format used by the program is explained in the following image. Should be easy to understand comparing left to right. I also give [several other examples](https://github.com/liyuanhe211/Energy_Diagram_Plotter_CDXML/tree/main/Examples):

![image](https://user-images.githubusercontent.com/18537705/160621422-05274905-5b1e-43b7-8cde-1ae80577d795.png)

The supported colors are expressed as follows (It's the [colors supported by Matplotlib](http://matplotlib.org/api/colors_api.html))

![image](https://user-images.githubusercontent.com/18537705/160621753-f0bad86a-5238-4f72-b7b0-c4554e0c0e0d.png)

The interface and the meaning of each option are described in the following figures (You can also hover on the options to see a tooltip):
The font of the figure is fixed to Arial (you can change it later in ChemDraw), and the canvas size and aspect ratio of the figure are determined by the window size

![image](https://user-images.githubusercontent.com/18537705/160618917-1544b242-6f63-45de-96e9-3892b549d445.png)

Here are some visuals for each options. Note that some of the following options are only reflected in the ChemDraw file, not in the preview view.

![image](https://user-images.githubusercontent.com/18537705/160621816-316c6304-1598-4a9c-83dc-666a071d2cb8.png)

![image](https://user-images.githubusercontent.com/18537705/160621848-128372f6-c51b-4674-bb1e-4a50adf370cb.png)

![image](https://user-images.githubusercontent.com/18537705/160621896-bbc03f1c-df92-43d2-8d60-0869156daec3.png)

![image](https://user-images.githubusercontent.com/18537705/160621889-39ee11a1-b997-48bd-b7ae-99887a75da99.png)

![image](https://user-images.githubusercontent.com/18537705/160621922-7f7f0526-fcb3-4df6-bbd2-3db4f1c1736c.png)

The last "Use temp file" option is to decide whether the program creates a file in its own Temp folder (the user needs to save it after viewing it), or directly generates a cdx file in the directory where Excel is located.

Other options not mentioned should be self-explanatory.

## Other issues

The temporary files of the program are stored in the program directory, so the program must have read and write permissions to this directory. It is not recommended to put the program in a path with permission restrictions such as Program Files, otherwise please give it administrator permissions.

The preview interface is not completely correct. If there is unexpected behavior, export the ChemDraw file first to see if it is normal.

The program's "Avoid text overlap" function is currently greedy, and may produce unreasonable results in particularly crowded situations. Just fine-tune it yourself (tip, hold down Shift and drag the text in ChemDraw to ensure it is horizontal or vertical). move).

![image](https://user-images.githubusercontent.com/18537705/160621958-1d139beb-878b-43b8-b882-4dbf56439ff0.png)

## For bug report

If you encounter some unexpected behavior, please try to run the [examples](https://github.com/liyuanhe211/Energy_Diagram_Plotter_CDXML/tree/main/Examples) first and see whether it's problem of your input or my program.

After that, you can give feedback by upload your Excel input file, a screenshot of the operation interface before crashing and the last display in the CMD window (if you what it is). If the output file is generated but there is a problem with the output file, please upload the output file as well. 

If there is a crash, please run the program through the "Draw_Energy_Diagram_XML_Debug.bat" file in the directory, so as to keep the error message before the program exits. 

