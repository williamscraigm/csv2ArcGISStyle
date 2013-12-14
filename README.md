# csv2ArcGISStyle

This is a utility for converting picture icons to marker symbols in an ArcGIS Style and ServerStyle.  The input is a csv file of a fixed format documenting the path of the input picture and basic descriptive information.

Supported picture types are PNG, JPEG, GIF, BMP and EMF. Use EMF if your source is vector artwork.  EMF picture markers will also be converted to Representation Markers in the ArcGIS style.

If your source artwork is SVG, these sources must be converted to EMF.  See the .bat file for a simple way to do this with Inkscape.

###CSV Format
The CSV file must contain the following fields:
- filePath: (string) the path to the input image.  It must be a complete path to the resource on disk (not a url)
- pointSize: (double) the point size of the marker symbol desired
- styleItemName: (string) the name you wish to list the symbol as in the style
- styleItemCategory: (string) the category you wish to assign to the symbol
- styleItemTags: (string) optional tags to aid with search of the symbol in ArcGIS 

A sample csv file with this information can be found in the samples folder