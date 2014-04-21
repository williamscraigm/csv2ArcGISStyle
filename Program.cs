/* 
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *    http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

using System;
using System.Collections.Generic;
using System.Text;
using ESRI.ArcGIS.ADF;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.Framework;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.DataSourcesOleDB;
using ESRI.ArcGIS.DataSourcesFile;
using ESRI.ArcGIS.Geodatabase;
using System.IO;
using System.Reflection;

namespace csv2ArcGISStyle
{
  
  class Program
  {
    private static LicenseInitializer m_AOLicenseInitializer = new csv2ArcGISStyle.LicenseInitializer();
  
    [STAThread()]
    static void Main(string[] args)
    {
      if (args.Length != 2)
      {
        Console.WriteLine("");
        Console.WriteLine("Invalid parameters specified");
        Console.WriteLine("Usage: csv2ArcGISStyle <csvFullPath> <styleFullPath>");
        Console.WriteLine("");
        Console.WriteLine("Example:");
        Console.WriteLine("csv2ArcGISStyle C:\\style\\myStyleSample.csv c:\\style\\myStyleSample.style");
        return;
      }

      //ESRI License Initializer generated code.
      System.Console.WriteLine("Initializing license...");
      if (!m_AOLicenseInitializer.InitializeApplication(new esriLicenseProductCode[] { esriLicenseProductCode.esriLicenseProductCodeAdvanced, esriLicenseProductCode.esriLicenseProductCodeStandard, esriLicenseProductCode.esriLicenseProductCodeBasic },
      new esriLicenseExtensionCode[] { }))
      {
        System.Console.WriteLine(m_AOLicenseInitializer.LicenseMessage());
        System.Console.WriteLine("This application could not initialize with the correct ArcGIS license and will shutdown.");
        m_AOLicenseInitializer.ShutdownApplication();
        return;
      }

      string csvPath = args[0];
      string stylePath = args[1];
      string serverStylePath = stylePath.Replace(".style", ".ServerStyle");

      File.Delete(stylePath); //delete the existing Style to start from scratch
      File.Delete(serverStylePath); //delete the existing Style to start from scratch

      //do the Desktop style
      Console.WriteLine("");
      Console.WriteLine("Creating the desktop style:");
      ImportCSV(csvPath, stylePath);
      ConvertVectorPicturesToRepresentationMarkers(stylePath);

      // do the Server styles + switch to disable if desired
      const bool DO_SERVER_STYLES = true;
      if (DO_SERVER_STYLES)
      {
        Console.WriteLine("");
        Console.WriteLine("Creating the server style:");
        ImportCSV(csvPath, serverStylePath);
        ConvertVectorPicturesToRepresentationMarkers(serverStylePath);
      }

      Console.WriteLine("");
      Console.WriteLine("Import complete, new styles " + stylePath + " and " + serverStylePath + " created");

      //ESRI License Initializer generated code.
      //Do not make any call to ArcObjects after ShutDownApplication()
      m_AOLicenseInitializer.ShutdownApplication();

    }

    static void ImportCSV(string csvPath, string stylePath)
    {
      //the expected field names for this util
      string filePath = "filePath";
      string pointSize = "pointSize";
      string styleItemName = "styleItemName";
      string styleItemCategory = "styleItemCategory";
      string styleItemTags = "styleItemTags";

      IStyleGallery styleGallery = GetStyleGallery(stylePath);
      
      IStyleGalleryStorage styleGalleryStorage = styleGallery as IStyleGalleryStorage;
      styleGalleryStorage.TargetFile = stylePath;

      ITable table = OpenCSVAsTable(csvPath);
      IRow row = null;
      int filePathIdx, styleItemNameIdx, pointSizeIdx, styleItemCategoryIdx, styleItemTagsIdx;
      IStyleGalleryItem3 styleGalleryItem = null;

      using (ComReleaser comReleaser = new ComReleaser())
      {
        // Create the cursor.
        ICursor cursor = table.Search(null, false);
        comReleaser.ManageLifetime(cursor);
        filePathIdx = cursor.FindField(filePath);
        styleItemNameIdx = cursor.FindField(styleItemName);
        pointSizeIdx = cursor.FindField(pointSize);
        styleItemCategoryIdx = cursor.FindField(styleItemCategory);
        styleItemTagsIdx = cursor.FindField(styleItemTags);

        int rowCount = 0;
        while ((row = cursor.NextRow()) != null)
        {
          rowCount++;

          try
          {
            String itemFilePath = (string)row.get_Value(filePathIdx);
            IPictureMarkerSymbol pictureMarkerSymbol = MakeMarkerSymbol(itemFilePath, Convert.ToDouble(row.get_Value(pointSizeIdx)));

            if (pictureMarkerSymbol == null)
            {
              Console.WriteLine("Image not found: Row: " + rowCount + ", Image: " + itemFilePath);
              continue;
            }

            styleGalleryItem = new StyleGalleryItemClass();
            styleGalleryItem.Item = pictureMarkerSymbol;
            styleGalleryItem.Name = row.get_Value(styleItemNameIdx) as string;
            styleGalleryItem.Category = row.get_Value(styleItemCategoryIdx) as string;
            object obj = row.get_Value(styleItemTagsIdx);

            // set a default tag, just in case tags field is empty
            string tags = styleGalleryItem.Category + ";" + styleGalleryItem.Name;

            // Make sure tags are set & a valid string type before converting to string            
            if (!((obj == null) || (obj is DBNull)))
                tags = obj as string;

            const int MAX_TAG_LENGTH = 255;

            int tagsLength = tags.Length;
            if (tagsLength > MAX_TAG_LENGTH)
            {
              // WORKAROUND if length > 255 (a hard limit), trim from the front (last most important)
              tags = tags.Substring(tagsLength - MAX_TAG_LENGTH, MAX_TAG_LENGTH);
            }

            // Note: Tag vector symbols for later (see: containsVectorTag) 
            // + switch to disable if desired
            const bool TAG_EMFS_WITH_VECTOR_TAG = true;
            if (TAG_EMFS_WITH_VECTOR_TAG && 
                (itemFilePath.Substring(itemFilePath.Length - 4) == ".emf") &&
                (!tags.ToUpper().Contains("VECTOR")))
            {
              tags = tags + ";vector";
            }

            styleGalleryItem.Tags = tags;

            //we want tags for search. If they weren't specified, use the default tags
            if (styleGalleryItem.Tags == "")
            {
                styleGalleryItem.Tags = styleGalleryItem.DefaultTags;
            }

            Console.WriteLine("Importing symbol " + rowCount + " : " + styleGalleryItem.Name);
            styleGallery.AddItem((IStyleGalleryItem)styleGalleryItem);
          }
          catch (Exception ex)
          {
            // Catch-all exception for row processing, just so 1 bad row doesn't abort the whole process
            Console.WriteLine("Skipping bad row: " + rowCount + ", Exception : " + ex.Message);
            continue;
          }
        }
      }

    }

    private static ITable OpenCSVAsTable(string csvPath)
    {
      IWorkspaceFactory wsFactory = new TextFileWorkspaceFactoryClass();
      FileInfo fileInfo = new FileInfo(csvPath);
      IWorkspace workspace = wsFactory.OpenFromFile(fileInfo.DirectoryName, 0);
      IFeatureWorkspace featWork = workspace as IFeatureWorkspace;
      ITable table = featWork.OpenTable(fileInfo.Name);
      return table;

    }

    private static IPictureMarkerSymbol MakeMarkerSymbol(string filePath, double size)
    {
      if (!System.IO.File.Exists(filePath))
      {
          Console.WriteLine("Source MarkerSymbol does NOT exist: " + filePath);
          return null;
      }

      bool isVector = (filePath.Substring(filePath.Length - 4) == ".emf");
      esriIPictureType picType = (isVector) ? esriIPictureType.esriIPictureEMF : esriIPictureType.esriIPicturePNG;
        
      IPictureMarkerSymbol pictureMarkerSymbol = new PictureMarkerSymbolClass();
      pictureMarkerSymbol.CreateMarkerSymbolFromFile(picType, filePath);
      pictureMarkerSymbol.Size = size;
      return pictureMarkerSymbol;
    }

    // A list of tags that identify a marker as a vector type
    private static List<string> vectorTags = new List<string>() { "EMF", "SVG", "VECTOR" };

    private static bool containsVectorTag(string tags)
    {
      string tagsUpper = tags.ToUpper();

      foreach (string vectorTag in vectorTags)
      {
        if (tagsUpper.Contains(vectorTag))
        {
          return true;
        }
      }

      return false;
    }

    private static void ConvertVectorPicturesToRepresentationMarkers(string stylePath) 
    {
      if (!System.IO.File.Exists(stylePath))
      {
          Console.WriteLine("StylePath does NOT exist: " + stylePath);
          return;
      }

      IStyleGallery styleGallery = GetStyleGallery(stylePath);
      IStyleGalleryStorage styleGalleryStorage = styleGallery as IStyleGalleryStorage;
      styleGalleryStorage.TargetFile = stylePath;
      IEnumStyleGalleryItem enumItems = styleGallery.get_Items("Marker Symbols", stylePath, "");

      IStyleGalleryItem3 originalItem = null;
      IStyleGalleryItem3 newItem = null;
      enumItems.Reset();

      originalItem = enumItems.Next() as IStyleGalleryItem3;
      while (originalItem != null)
      {
        IPictureMarkerSymbol pictureMS = originalItem.Item as IPictureMarkerSymbol;
        if (pictureMS != null)
        {
          // check if it requires conversion to rep marker? 
          // TODO: perhaps make this an option if user doesn't need rep's
          if (containsVectorTag(originalItem.Tags)) 
          {
            newItem = ConvertMarkerItemToRep(originalItem);
            Console.WriteLine("Converting symbol " + newItem.Name + " to a representation marker");
            styleGallery.AddItem(newItem);
          }
          originalItem = enumItems.Next() as IStyleGalleryItem3;
        }
      }
      
    }

    private static IStyleGalleryItem3 ConvertMarkerItemToRep(IStyleGalleryItem3 inputItem)
    {
      IMarkerSymbol markerSymbol = inputItem.Item as IMarkerSymbol;
      IRepresentationRule repRule = new RepresentationRuleClass();
      IRepresentationRuleInit repRuleInit = repRule as IRepresentationRuleInit;

      repRuleInit.InitWithSymbol((ISymbol)markerSymbol); //initialize the rep rule with the marker
      IRepresentationGraphics representationGraphics = new RepresentationMarkerClass();

      IGraphicAttributes graphicAttributes = null;
      IRepresentationGraphics tempMarkerGraphics = null;
      IGeometry tempGraphicGeometry = null;
      IRepresentationRule tempRule = null;

      //only pull the markers out.
      for (int i = 0; i < repRule.LayerCount; i++)
      {

        graphicAttributes = repRule.get_Layer(i) as IGraphicAttributes;
        tempMarkerGraphics = graphicAttributes.get_Value((int)esriGraphicAttribute.esriGAMarker) as IRepresentationGraphics;

        tempMarkerGraphics.Reset();
        tempMarkerGraphics.Next(out tempGraphicGeometry, out tempRule);

        while (tempRule != null && tempGraphicGeometry != null)
        {
          representationGraphics.Add(tempGraphicGeometry, tempRule);
          tempGraphicGeometry = null;
          tempRule = null;
          tempMarkerGraphics.Next(out tempGraphicGeometry, out tempRule);
        }
      }

      IStyleGalleryItem3 newMarkerStyleGalleryItem = new ServerStyleGalleryItemClass();
      newMarkerStyleGalleryItem.Item = representationGraphics;
      newMarkerStyleGalleryItem.Name = inputItem.Name;
      newMarkerStyleGalleryItem.Category = inputItem.Category;
      newMarkerStyleGalleryItem.Tags = inputItem.Tags.Replace(";emf", ""); //strip emf from the tags

      return newMarkerStyleGalleryItem;
    }

    static IStyleGallery GetStyleGallery(string stylePath)
    {
      IStyleGallery styleGallery = null;
      String tempPath = stylePath.ToLower();
      if (tempPath.Contains(".serverstyle"))
      {
        styleGallery = new ServerStyleGalleryClass();
      }
      else
      {
        styleGallery = new StyleGalleryClass();
      }
      return styleGallery;
    }

  }

}