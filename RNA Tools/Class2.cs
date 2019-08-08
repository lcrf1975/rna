// Geocode API Reference
// https://github.com/chadly/Geocoding.net
using Geocoding;
using Geocoding.Google;
using Geocoding.Microsoft;

// EEPlus - Excel API
// http://epplus.codeplex.com/

// Sharp KML
// https://sharpkml.codeplex.com/
using SharpKml.Base;
using SharpKml.Dom;
using SharpKml.Engine;

using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace AUX
{
    class Services
    {
        // Class to support the KML creation process
        public class Location
        {
            public string Name;
            public double Lat;
            public double Lon;
            public string Address;
            public string Quality;
        }

        // Geocodes an address using the Bing Maps engine
        public static Location SearchBing(string add)
        {
            try
            {
                // Calls the Werservice
                BingMapsGeocoder geocoder = new BingMapsGeocoder(ConfigurationManager.AppSettings["BingApiKey"]);
                IEnumerable<Address> addresses = geocoder.Geocode(add);

                // Selects the firt result
                BingAddress b = (BingAddress)addresses.FirstOrDefault();

                Location r = new Location();

                r.Lat = addresses.FirstOrDefault().Coordinates.Latitude;
                r.Lon = addresses.FirstOrDefault().Coordinates.Longitude;
                r.Quality = b.Confidence.ToString();
                r.Address = addresses.FirstOrDefault().FormattedAddress;

                return r;
            }
            catch (Exception Ex)
            {
                throw new Exception(Ex.Message);
            }
        }

        // Geocodes an address using the Bing Maps engine
        public static Location SearchGoogle(string add)
        {
            try
            {
                // Calls the webservice
                GoogleGeocoder geocoder = new GoogleGeocoder() { ApiKey = ConfigurationManager.AppSettings["GoogleApiKey"] };
                IEnumerable<Address> addresses = geocoder.Geocode(add);

                // Selects the firt result
                GoogleAddress g = (GoogleAddress)addresses.FirstOrDefault();

                Location r = new Location();

                r.Lat = addresses.FirstOrDefault().Coordinates.Latitude;
                r.Lon = addresses.FirstOrDefault().Coordinates.Longitude;
                r.Quality = g.LocationType.ToString();
                r.Address = addresses.FirstOrDefault().FormattedAddress;

                return r;
            }
            catch (Exception Ex)
            {
                throw new Exception(Ex.Message);
            }
        }

        // Creates a KML file from an array of locations
        public static void CreateKML(string fileName, Location[] locations)
        {
            Document document = new Document();

            foreach (Location l in locations)
            {
                // Creates a new Placemark
                Point point = new Point();
                point.Coordinate = new Vector(l.Lat, l.Lon);
                Placemark placemark = new Placemark();
                placemark.Name = l.Name;
                placemark.Address = l.Address;
                placemark.Geometry = point;

                // Stores the Placemark in a the document
                document.AddFeature(placemark);
            }

            // Converts the documento to KML format
            Kml root = new Kml();
            root.Feature = document;
            KmlFile kml = KmlFile.Create(root, false);

            try
            {
                // Delete the old file if it exits
                if (File.Exists(@fileName))
                {
                    File.Delete(@fileName);
                }

                // Creates the file
                using (FileStream stream = File.OpenWrite(fileName))
                {
                    // Saves the stream to the selected path
                    kml.Save(stream);
                    MessageBox.Show("File " + fileName + " created.");
                }
            }
            catch (Exception Ex)
            {
                throw new Exception(Ex.Message);
            }
        }

        // Converts an Excel xlsx file to a ListView
        public static ListView ExcelToListView(string path)
        {
            try
            {
                using (var pck = new OfficeOpenXml.ExcelPackage())
                {
                    // Opens the xlxs file
                    using (var stream = File.OpenRead(path))
                    {
                        pck.Load(stream);
                    }

                    ListView list = new ListView(); // Creates a new listview to store the data

                    var ws = pck.Workbook.Worksheets.First(); // Select the first worksheet

                    // Creates the listview columns
                    foreach (var col in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                    {
                        list.Columns.Add(col.Text);
                    }

                    // Loads the worksheet data into the list
                    for (int rowNum = 2; rowNum <= ws.Dimension.End.Row; rowNum++)
                    {
                        ListViewItem listRow = new ListViewItem();

                        for (int ColNum = 1; ColNum <= ws.Dimension.End.Column; ColNum++)
                        {
                            listRow.SubItems.Add(ws.Cells[rowNum, ColNum, rowNum, ColNum].Text);
                        }

                        list.Items.Add(listRow);
                    }

                    return list;
                }
            }
            catch (Exception Ex)
            {
                throw new Exception(Ex.Message);
            }
        }

        // Auto resize a listview object - Very Usefull
        public static void autoResizeColumns(ListView lv, int factor)
        {
            // Auto size by column width
            lv.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);

            // Check and resize by content width 
            ListView.ColumnHeaderCollection cc = lv.Columns;
            for (int i = 0; i < cc.Count; i++)
            {
                int colWidth = TextRenderer.MeasureText(cc[i].Text, lv.Font).Width + factor;
                if (colWidth > cc[i].Width)
                {
                    cc[i].Width = colWidth;
                }
            }
        }
    }
}