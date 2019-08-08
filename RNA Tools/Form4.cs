// RNA Web Service Reference
using RNA_Tools.Apex;

using System;
using System.Linq;
using System.Windows.Forms;

namespace RNA_Tools
{
    public partial class Form4 : Form
    {
        public string Route_Name;
        public Form4(long rsession, long rkey, string rname)
        {
            InitializeComponent();

            // Change the form name
            Route_Name = rname;
            this.Text = "Route [" + Route_Name + "] Stops:";

            // Loads the Stops data into the list
            getStops(rsession, rkey);

            // Auto Size the listview columns
            AUX.Services.autoResizeColumns(listView1, 20);
        }

        // Loads all Stops from the selected Route
        private void getStops(long route_session, long route_key)
        {
            try
            {
                // Clear the Stops list
                listView1.Items.Clear();

                // Loads all Route data from the Route Session
                Route[] Routes = RNA.WebServices.RetrieveRoutes(route_session);

                // Search the select Route data in the Route array
                foreach (Route route in Routes)
                {
                    if (route.EntityKey == route_key)
                    {
                        if (route.NumberOfServiceableStops > 0)
                        {
                            // Loads Origin Depot data into the list
                            string[] dep = RNA.WebServices.SearchDepot((long)route.OriginDepotEntityKey);

                            listView1.Items.Add(new ListViewItem(new string[]
                            {
                                        "0",
                                        route.OriginDepotEntityKey.ToString(),
                                        "Origin",
                                        dep[0], // Name
                                        dep[1], // Address
                                        dep[2], // Lat
                                        dep[3]  // Lon
                            }));

                            for (int i = 0; i < route.NumberOfServiceableStops; i++)
                            {
                                if (route.Stops[i] is ServiceableStop)
                                {
                                    // Converts Lat & Lon to Degrees
                                    double lat = route.Stops[i].Coordinate.Latitude * 0.000001;
                                    double lon = route.Stops[i].Coordinate.Longitude * 0.000001;

                                    ServiceableStop serviceableStop = (ServiceableStop)route.Stops[i];

                                    string stop_type = serviceableStop.Actions[0].StopActionType_Type;

                                    // Searches for a different stop type
                                    if (serviceableStop.Actions.Count() > 1)
                                    {
                                        for (int c = 1; c < serviceableStop.Actions.Count(); c++)
                                        {
                                            if (stop_type != serviceableStop.Actions[c].StopActionType_Type)
                                            {
                                                stop_type = stop_type + "&" + serviceableStop.Actions[c].StopActionType_Type;
                                                c = serviceableStop.Actions.Count();
                                            }
                                        }
                                    }

                                    // Loads Stops data into the list
                                    listView1.Items.Add(new ListViewItem(new string[]
                                    {
                                        (i+1).ToString(),
                                        route.Stops[i].EntityKey.ToString(),
                                        stop_type,
                                        serviceableStop.ServiceLocationDescription,
                                        serviceableStop.ServiceLocationAddress.AddressLine1 + ", "
                                            + serviceableStop.ServiceLocationAddress.Locality.AdminDivision3 + ", " // City
                                            + serviceableStop.ServiceLocationAddress.Locality.AdminDivision1 + ", " // State
                                            + serviceableStop.ServiceLocationAddress.Locality.PostalCode,
                                        lat.ToString(),
                                        lon.ToString()
                                    }));
                                }
                            }

                            // Loads Destination Depot data into the list
                            dep = RNA.WebServices.SearchDepot((long)route.DestinationDepotEntityKey);

                            listView1.Items.Add(new ListViewItem(new string[]
                            {
                                        listView1.Items.Count.ToString(),
                                        route.DestinationDepotEntityKey.ToString(),
                                        "Destination",
                                        dep[0], // Name
                                        dep[1], // Address
                                        dep[2], // Lat
                                        dep[3]  // Lon
                            }));
                        }
                        else
                        {
                            MessageBox.Show("No stops found.", "Warning");
                        }

                        // Exit loop
                        break;
                    }
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Warning");
            }
        }

        // Exports the stops (Lat & Lon) to a KML file
        private void button1_Click(object sender, EventArgs e)
        {
            // Checks if the list is empty
            if (listView1.Items.Count > 0)
            {
                int c = 0;
                AUX.Services.Location[] Loc = new AUX.Services.Location[listView1.Items.Count];

                // Loads the stop list into the Location class
                foreach (ListViewItem item in listView1.Items)
                {
                    Loc[c] = new AUX.Services.Location();
                    Loc[c].Name = item.SubItems[3].Text;
                    Loc[c].Address = item.SubItems[4].Text;
                    Loc[c].Lat = Convert.ToDouble(item.SubItems[5].Text);
                    Loc[c].Lon = Convert.ToDouble(item.SubItems[6].Text);
                    c++;
                }

                // Displays a SaveFileDialog
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.InitialDirectory = Convert.ToString(Environment.SpecialFolder.MyDocuments);
                saveFileDialog1.FileName = "Route " + Route_Name + ".kml";
                saveFileDialog1.DefaultExt = "*.kml";
                saveFileDialog1.Filter = "KML File|*.kml";
                saveFileDialog1.Title = "Export Stops to a KML File";

                // If the file name returns OK
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    // Creates the KML file
                    AUX.Services.CreateKML(saveFileDialog1.FileName, Loc);
                }
            }
            else
            {
                MessageBox.Show("No stops to export.", "Warning");
            }
        }

        // Generates the Travel Path from PoitnA to B - Under Working...
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                button2.Text = "Working...";

                int c = 0;

                // Creates a coordinate array to store the Path points
                Coordinate[] Coord = new Coordinate[listView1.Items.Count];

                // Loads the stop list into the Coordinate array
                foreach (ListViewItem item in listView1.Items)
                {
                    Coord[c] = new Coordinate();
                    Coord[c].Latitude = (int)(Convert.ToDouble(item.SubItems[5].Text) / 0.000001); // Convert from Double to Int
                    Coord[c].Longitude = (int)(Convert.ToDouble(item.SubItems[6].Text) / 0.000001);
                    c++;
                }

                // Call the WebService
                TravelPathResult r = RNA.WebServices.TravelPath(Coord);

                // Loads the result into the Location Class
                c = 0;
                AUX.Services.Location[] Loc = new AUX.Services.Location[r.Path.Points.Count()];
              
                // Loads the returned points into the Location class required to build the KML file
                foreach (Coordinate p in r.Path.Points) // Load all points from the Path
                {
                    Loc[c] = new AUX.Services.Location();
                    Loc[c].Lat = p.Latitude * 0.000001; // Convert from Int to Double
                    Loc[c].Lon = p.Longitude * 0.000001;
                    c++;
                }

                button2.Text = "Travel Path";

                // Displays a SaveFileDialog
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.InitialDirectory = Convert.ToString(Environment.SpecialFolder.MyDocuments);
                saveFileDialog1.FileName = "Route " + Route_Name + ".kml";
                saveFileDialog1.DefaultExt = "*.kml";
                saveFileDialog1.Filter = "KML File|*.kml";
                saveFileDialog1.Title = "Export Travel Path to a KML File";

                // If the file name returns OK
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    // Creates the KML file
                    AUX.Services.CreateKML(saveFileDialog1.FileName, Loc);
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Error");
            }
        }
    }
}
