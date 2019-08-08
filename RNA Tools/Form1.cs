using Microsoft.VisualBasic;

// RNA Web Service Reference
using RNA_Tools.Apex;

using System;
using System.Configuration;
using System.Diagnostics;
using System.Windows.Forms;

namespace RNA_Tools
{
    public partial class Form1 : Form
    {
        public long BU_EntityKey; // Caches the BU EntityKey
        public string BU_Name; // Caches the BU Name
        public long Login_EntityKey; // Caches the logged used EntityKey
        public long Region_EntityKey; // Caches the selected Region EntityKey
        public long RouteSession_EntityKey; // Caches the selected Route Session EntityKey
        public string RouteSession_Date; // Caches the selected Route Session start date
        public long Route_EntityKey; // Caches the selected Route EntityKey

        public Form1()
        {
            InitializeComponent();

            // Initializes the global variables
            BU_EntityKey = 0;
            Region_EntityKey = 0;
            RouteSession_EntityKey = 0;
            Login_EntityKey = 0;

            // Clears the tool strip labels of status bar
            toolStripStatusLabel2.Text = "";
            toolStripStatusLabel3.Text = "";
            toolStripStatusLabel4.Text = "";
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Displays the About Form.
            Form2 frm = new Form2();
            frm.ShowDialog();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Displays a message box asking if the user want to exit the application.
            if (MessageBox.Show("Do you want to exit?", "Message",
                MessageBoxButtons.YesNo)
                == DialogResult.Yes)
            {
                Application.Exit();
            }
        }

        // Webservices connection procedure (Login)
        private void connectToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Connecting...";

            try
            {
                // Login procedure call
                string[] userEntityKey = RNA.WebServices.Login();

                // Checks if a valid user key was returned
                Login_EntityKey = Convert.ToInt32(userEntityKey[0]); // Caches the logged used EntityKey

                if (Login_EntityKey > 0)
                {
                    // Checks if there are BU for the user
                    if (getBusinessUnits() > 0)
                    {
                        toolStripStatusLabel2.Text = BU_Name;

                        toolStripStatusLabel1.Text = "Getting Regions...";

                        // Retrieves all BU regions
                        getRegions();

                        // Auto Size the Regions listview columns
                        AUX.Services.autoResizeColumns(listView1, 20);

                        // Update the form with the Connected status
                        connectToolStripMenuItem.Checked = true;
                        toolStripStatusLabel1.Text = "Connected.";
                    }
                    else
                    {
                        MessageBox.Show("No Business Units found.", "Error");
                    }
                }
                else
                {
                    connectToolStripMenuItem.Checked = false;
                    toolStripStatusLabel1.Text = "Not Connected.";
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Error");

                // Forces the application termination becouse of problems with the webservice connection
                Application.Exit();
            }
        }

        private void optionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Displays the Options Form.
            Form3 frm = new Form3();
            frm.ShowDialog();
        }

        // Loads all BU associated with the user
        private int getBusinessUnits()
        {
            try
            {
                // Loads all BU into the BU array
                BusinessUnit[] bus = RNA.WebServices.RetrieveBusinessUnits();
                if (bus.Length > 1)
                {
                    MessageBox.Show("RNA Tool automatic selects the first Business Unit.", "Warning");
                }

                // Stores the First Bussiness key and name
                BU_EntityKey = bus[0].EntityKey;
                BU_Name = bus[0].Identifier;

                // Stores the BU country. Needed to geocode address
                textBox5.Text = bus[0].Defaults.CountryISO3Abbr; // Country

                return bus.Length;
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Error");
                return 0;
            }

        }

        // Retrives all Regions
        private void getRegions()
        {
            try
            {
                // Clears the Region list
                listView1.Items.Clear();
                label8.Text = "Count: 0";

                // Loads all Regions into the Region array
                Region[] regions = RNA.WebServices.RetrieveRegions();

                // Loads the Region data into a list
                foreach (Region region in regions)
                {
                    listView1.Items.Add(new ListViewItem(new string[] 
                    {
                        region.EntityKey.ToString(),
                        region.Identifier,
                        region.Description
                    }));
                }

                label1.Text = "Count: " + listView1.Items.Count.ToString(); ;
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Error");
            }

        }

        // Loads the Webservice URL data
        private void getURLs()
        {
            try
            {
                // Clears the URL list
                listView2.Items.Clear();

                // Loads the URL data into a list and create the URL Webservice objects
                foreach (String s in RNA.WebServices.RetrieveUrls(BU_EntityKey, Int32.Parse(listView1.SelectedItems[0].SubItems[0].Text)))
                {
                    listView2.Items.Add(new ListViewItem(new string[] { s, "Ready" }));
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Error");
            }
        }

        // Loads the Routing Session data
        private void getRouteSessions()
        {
            try
            {
                // Clears the Route Sessions List
                listView4.Items.Clear();
                label8.Text = "Count: 0";

                // Loads the Route Session data into an array
                DailyRoutingSession[] routeSessions = RNA.WebServices.RetrieveRouteSessions();

                // Loads the Route Session array data into a list
                foreach (DailyRoutingSession session in routeSessions)
                {
                    listView4.Items.Add(new ListViewItem(new string[]
                    {
                        session.EntityKey.ToString(),
                        session.StartDate.ToString(),
                        session.Description
                    }));
                }

                label8.Text = "Count: " + listView4.Items.Count.ToString(); ;
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Error");
            }
        }

        // Loads the Route data
        private void getRoutes()
        {
            try
            {
                // Clears the Route list
                listView5.Items.Clear();
                label2.Text = "Count: 0";

                // Loads the Route data into an array
                Route[] Routes = RNA.WebServices.RetrieveRoutes(RouteSession_EntityKey);

                // Loads the Route array data into a list
                foreach (Route route in Routes)
                {
                    listView5.Items.Add(new ListViewItem(new string[]
                    {
                        route.EntityKey.ToString(),
                        route.Identifier + " - "
                        + route.Description,
                        route.RoutePhase_Phase,
                        String.Format("{0:00}", route.TotalTime.Hours) + ":"
                        + String.Format("{0:00}", route.TotalTime.Minutes) + ":"
                        + String.Format("{0:00}", route.TotalTime.Seconds),
                        route.NumberOfServiceableStops.ToString(),
                        String.Format("{0:0.00}", route.TotalDistance.Value) + " "
                        + route.TotalDistance.DistanceUnit_Unit,
                        String.Format("{0:0.00}", route.Costs.TotalCost),
                        RNA.WebServices.SearchDepot(route.OriginDepotEntityKey)[0] // Name
                    }));
                }

                label2.Text = "Count: " + listView5.Items.Count.ToString(); ;

            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Warning");
            }
        }

        // Loads the Service Loation data
        private void getServiceLocations()
        {
            try
            {
                // Clears the Service Location List
                listView6.Items.Clear();
                label9.Text = "Count: 0";

                // Loads the Service Location data into an array
                ServiceLocation[] serviceLocations = RNA.WebServices.RetrieveServiceLocation();

                // Loads the Service Location array data into a list
                foreach (ServiceLocation slocation in serviceLocations)
                {
                    // Bad service location filter
                    if (slocation.IsDeleted == false &&
                        slocation.Description.Length >0
                        && slocation.Address.AddressLine1.Length > 0)
                    {
                        listView6.Items.Add(new ListViewItem(new string[]
                        {
                            slocation.EntityKey.ToString(),
                            slocation.Identifier,
                            slocation.Description,
                            slocation.Address.AddressLine1 + ", " +
                            slocation.Address.Locality.AdminDivision2 + " - "+ // City
                            slocation.Address.Locality.AdminDivision1 + " "+ // State
                            slocation.Address.Locality.PostalCode
                        }));
                    }
                }

                label9.Text = "Count: " + listView6.Items.Count.ToString(); ;

            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Warning");
            }
        }
 
        // Loads the Serive Orders data
        private void getServiceOrders()
        {
            try
            {
                // Clears the Service Orders List
                listView7.Items.Clear();
                label10.Text = "";

                // Loads the Service Orders data into an array
                Order[] orders = RNA.WebServices.RetrieveServiceOrders(RouteSession_EntityKey);

                // Loads the Service Orders array data into a list
                foreach (Order order in orders)
                {
                    listView7.Items.Add(new ListViewItem(new string[]
                    {
                        order.EntityKey.ToString(),
                        order.Tasks[0].LocationDescription,
                        order.Tasks[0].LocationAddress.AddressLine1 + ", "
                        + order.Tasks[0].LocationAddress.Locality.AdminDivision3 + " - "
                        + order.Tasks[0].LocationAddress.Locality.AdminDivision1 + " "
                        + order.Tasks[0].LocationAddress.Locality.PostalCode,
                        order.OrderType_Type,
                        order.BeginDate,
                        order.EndDate,
                        order.OrderState_State
                    }));
                }

                label10.Text = "Count: " + listView7.Items.Count.ToString(); ;

            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Warning");
            }
        }
  
        // Loads data from the selected Region
        private void button1_Click(object sender, EventArgs e)
        {
            // Checks the Region list has items
            if (listView1.Items.Count > 0)
            {
                // Checks if there is a selected item
                if (listView1.SelectedItems.Count > 0)
                {
                    // Stores the Region key for future usage
                    Region_EntityKey = Convert.ToInt32(listView1.SelectedItems[0].SubItems[0].Text);

                    button1.Text = "Working...";

                    // This can take a lond time...
                    getURLs(); //Loads Webservice URL data
                    getRouteSessions(); // Loads Route Sessions data
                    getServiceLocations(); // Loads Service Locations data

                    // Auto Size the Services listview columns
                    AUX.Services.autoResizeColumns(listView2, 20);

                    // Auto Size the Services listview columns
                    AUX.Services.autoResizeColumns(listView4, 20);

                    // Auto Size the Locations listview columns
                    AUX.Services.autoResizeColumns(listView6, 20);

                    button1.Text = "Select";

                    toolStripStatusLabel3.Text = listView1.SelectedItems[0].SubItems[1].Text;
                    tabControl1.SelectedTab = tabPage4; // Route Sessions
                }
                else
                {
                    MessageBox.Show("No Region selected.", "Warning");
                }
            }
            else
            {
                MessageBox.Show("Connect to RNA Web Service first <Ctrl+C>.", "Warning");
            }
        }

        // Geocodes an serive location address
        private void button2_Click(object sender, EventArgs e)
        {
            if (listView2.Items.Count > 0)
            {
                // Clears Geocode address list
                listView3.Items.Clear();
 
                try
                {
                    button2.Text = "Working...";

                    // Loads the address data into the object
                    Address address = new Address
                    {
                        AddressLine1 = textBox1.Text.Trim(),
                        AddressLine2 = "",
                        Locality = new Locality
                        {
                            AdminDivision1 = textBox3.Text.Trim(), // State
                            AdminDivision2 = textBox2.Text.Trim(), // City
                            CountryISO3Abbr = textBox5.Text, // Country
                            PostalCode = textBox4.Text.Trim()
                        }
                    };

                    // RNA Geocoding
                    int c = 0;
                    char[] delimiterChars = { '|' };
                    foreach (String s in RNA.WebServices.Geocode(address))
                    {
                        c++;
                        string[] r = s.Split(delimiterChars);
                        listView3.Items.Add(new ListViewItem(new string[]
                        {
                            c.ToString(),
                            "RNA Maps",
                            r[0],
                            r[1],
                            r[2],
                            r[3],
                            "-"
                        }));
                    }

                    // Bing Geocoding
                    if (checkBox1.Checked)
                    {   
                        c++;
                        AUX.Services.Location b = AUX.Services.SearchBing(textBox1.Text + " "
                            + textBox2.Text + " "
                            + textBox3.Text + " "
                            + textBox4.Text + " "
                            + textBox5.Text);
                        listView3.Items.Add(new ListViewItem(new string[]
                        {
                            c.ToString(),
                            "Bing Maps",
                            b.Lat.ToString(),
                            b.Lon.ToString(),
                            "-",
                            b.Quality,
                            b.Address
                        }));
                    }

                    // Google Geocoding
                    if (checkBox2.Checked)
                    {
                        c++;
                        AUX.Services.Location g = AUX.Services.SearchGoogle(textBox1.Text + " "
                            + textBox2.Text + " "
                            + textBox3.Text + " "
                            + textBox4.Text + " "
                            + textBox5.Text);
                        listView3.Items.Add(new ListViewItem(new string[]
                        {
                            c.ToString(),
                            "Google Maps",
                            g.Lat.ToString(),
                            g.Lon.ToString(),
                            "-",
                            g.Quality,
                            g.Address
                        }));
                    }


                    // Auto Size the Geocode Result listview columns
                    AUX.Services.autoResizeColumns(listView3, 20);

                    button2.Text = "Search";
                }
                catch (Exception Ex)
                {
                    MessageBox.Show(Ex.Message, "Error");
                }
            }
            else
            {
                MessageBox.Show("No services loaded. Select one region first.", "Warning");
            }
        }

        // Loads the Route Session data
        private void button3_Click(object sender, EventArgs e)
        {
            if (listView4.SelectedItems.Count > 0)
            {
                // Caches the Route Session data
                RouteSession_EntityKey = Convert.ToInt32(listView4.SelectedItems[0].SubItems[0].Text);
                RouteSession_Date = listView4.SelectedItems[0].SubItems[1].Text;

                button3.Text = "Working...";

                // From the selected Route Sessin
                getRoutes(); // Loads the Routes
                getServiceOrders(); // Loads the Service Orders

                // Auto Size the Routes listview columns
                AUX.Services.autoResizeColumns(listView5, 20);

                // Auto Size the Orders listview columns
                AUX.Services.autoResizeColumns(listView7, 20);

                toolStripStatusLabel4.Text = listView4.SelectedItems[0].SubItems[1].Text;

                button3.Text = "Select";

                // Checks if there are Routes in the list
                if (listView5.Items.Count > 0)
                {
                    tabControl1.SelectedTab = tabPage5; // Routes
                }
            }
            else
            {
                MessageBox.Show("No Route Session selected.", "Warning");
            }
        }

        // Shows the Stop information from the selected Route
        private void button4_Click(object sender, EventArgs e)
        {
            if (listView5.SelectedItems.Count > 0)
            {
                // Stores the selected Route key
                Route_EntityKey = Convert.ToInt32(listView5.SelectedItems[0].SubItems[0].Text);

                button4.Text = "Working...";

                // Displays the About Form.
                Form4 frm = new Form4(RouteSession_EntityKey, Route_EntityKey, listView5.SelectedItems[0].SubItems[1].Text);
                frm.ShowDialog();

                // Removes the route selection to force the user select again
                listView5.SelectedItems.Clear();

                button4.Text = "Select";
            }
            else
            {
                MessageBox.Show("No Route selected.", "Warning");
            }
        }

        // Forces the from to refresh if the tool strip text change
        private void toolStripStatusLabel1_TextChanged(object sender, EventArgs e)
        {
            this.Update();
        }

        // Shows the found geocode in the Google Maps
        private void button7_Click(object sender, EventArgs e)
        {
            if (listView3.Items.Count > 0)
            {
                // RNA Coordinates
                ListViewItem R = listView3.FindItemWithText("RNA Maps");
                string R_Lat = R.SubItems[2].Text.Replace(",", ".");
                string R_Lon = R.SubItems[3].Text.Replace(",", ".");
                string R_Location = R_Lat + "," + R_Lon;

                string B_Location = "";

                // Bing Coordinates
                if (checkBox1.Checked)
                {
                    ListViewItem B = listView3.FindItemWithText("Bing Maps");
                    string B_Lat = B.SubItems[2].Text.Replace(",", ".");
                    string B_Lon = B.SubItems[3].Text.Replace(",", ".");
                    B_Location = "&markers=color:blue|label:B|" + B_Lat + "," + B_Lon;
                }

                string G_Location = "";
           
                // Goolge Coordinates
                if (checkBox2.Checked)
                {
                    ListViewItem G = listView3.FindItemWithText("Google Maps");
                    string G_Lat = G.SubItems[2].Text.Replace(",", ".");
                    string G_Lon = G.SubItems[3].Text.Replace(",", ".");
                    G_Location = "&markers=color:green|label:G|" + G_Lat + "," + G_Lon;
                }

                // Calls the IExplorer browser with the required parameters
                try
                {
                    Process ExternalProcess = new Process();
                    ExternalProcess.StartInfo.FileName = "iexplore.exe";
                    ExternalProcess.StartInfo.Arguments = "https://maps.google.com/maps/api/staticmap?center="
                        + R_Location + "&zoom=17&size=640x640&scale=2"
                        + "&markers=color:red|label:R|" + R_Location
                        + B_Location
                        + G_Location
                        + "&key=" + ConfigurationManager.AppSettings["GoogleApiKey"];
                    ExternalProcess.StartInfo.WindowStyle = ProcessWindowStyle.Maximized;
                    ExternalProcess.Start();
                    ExternalProcess.WaitForExit();
                }
                catch (Exception Ex)
                {
                    MessageBox.Show(Ex.Message, "Error");
                }
            }
            else
            {
                MessageBox.Show("No location coordinates to view.", "Warning");
            }
        }

        // Creates a new route session
        private void button8_Click(object sender, EventArgs e)
        {
            if (Region_EntityKey > 0)
            {
                // Capture the next day date to be use as default name
                DateTime now = DateTime.Now;

                if (now.DayOfWeek == System.DayOfWeek.Friday)
                {
                    now = now.AddDays(3); // Next week
                }
                else
                {
                    now = now.AddDays(1); // Next day
                }

                now.AddDays(1); // Next day

                string r = String.Format("{0:0000}", now.Year) + "-"
                    + String.Format("{0:00}", now.Month) + "-"
                    + String.Format("{0:00}", (now.Day));

                // Visual Basic Interop - No Input Call in C#
                r = Interaction.InputBox("Inform or confirm the new route session date:", "New Route Session", r, -1, -1);

                if (r.Trim().Length > 0) // OK Button
                {
                    // Converts the result string into a system date time
                    DateTime d = Convert.ToDateTime(r);

                    // Create the Dauly Rouse Session object
                    DailyRoutingSession newRouteSession = new DailyRoutingSession();

                    // Load the Route Session information into the object
                    newRouteSession.Action = ActionType.Add;
                    newRouteSession.CreatedBy = ConfigurationManager.AppSettings["Login"];
                    newRouteSession.Description = "Created by RNA WebSerives Tool on " + d.ToShortDateString();
                    newRouteSession.NumberOfTimeUnits = 1;
                    newRouteSession.RegionEntityKey = Region_EntityKey;
                    newRouteSession.SessionMode_Mode = Enum.GetName(typeof(SessionMode), SessionMode.Operational);
                    newRouteSession.StartDate = r;
                    newRouteSession.TimeUnit_TimeUnitType = Enum.GetName(typeof(TimeUnit), TimeUnit.Day);

                    try
                    {
                        // Create the new Route Session
                        long entity_key = 0;
                        entity_key = RNA.WebServices.NewRoutingSession(newRouteSession);

                        // Check if the Route Session was created.
                        if (entity_key > 0)
                        {
                            getRouteSessions(); // Update Route Sessions data

                            MessageBox.Show("New Route Session " + r + " created. Entity Key " + entity_key.ToString());
                        }
                    }
                    catch (Exception Ex)
                    {
                        MessageBox.Show(Ex.Message, "Error");
                    }
                }
            }
            else
            {
                MessageBox.Show("Select one Region first.", "Warning");
            }
        }

        // Import Orders
        private void button6_Click(object sender, EventArgs e)
        {
            if (toolStripStatusLabel4.Text.Length > 0)
            {
                // Displays a OpenFileDialog
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.InitialDirectory = Convert.ToString(Environment.SpecialFolder.MyDocuments);
                openFileDialog1.DefaultExt = "*.xlsx";
                openFileDialog1.Filter = "Excel File|*.xlsx";
                openFileDialog1.Title = "Select the Excel File to Import";

                // If the file name returns OK
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    // Displays the Import Form.
                    Form5 frm = new Form5("Import Order File: " + openFileDialog1.FileName,
                        AUX.Services.ExcelToListView(openFileDialog1.FileName), 1, // 1 = Orders
                        Login_EntityKey, BU_EntityKey, Region_EntityKey, RouteSession_EntityKey, RouteSession_Date, listView6);
                    frm.ShowDialog();
                }
            }
            else
            {
                MessageBox.Show("No Route Session loaded.", "Warning");
            }
        }

        // Import Service Locations
        private void button5_Click(object sender, EventArgs e)
        {
            if (toolStripStatusLabel3.Text.Length > 0)
            {
                // Displays a OpenFileDialog
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.InitialDirectory = Convert.ToString(Environment.SpecialFolder.MyDocuments);
                openFileDialog1.DefaultExt = "*.xlsx";
                openFileDialog1.Filter = "Excel File|*.xlsx";
                openFileDialog1.Title = "Select the Excel File to Import";

                // If the file name returns OK
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    // Displays the Import Form.
                    Form5 frm = new Form5("Import Location File: " + openFileDialog1.FileName,
                        AUX.Services.ExcelToListView(openFileDialog1.FileName), 2, // 2 = Locations
                        Login_EntityKey, BU_EntityKey, Region_EntityKey, RouteSession_EntityKey, RouteSession_Date, listView6);
                    frm.ShowDialog();
                }
            }
            else
            {
                MessageBox.Show("No Region loaded.", "Warning");
            }
        }
    }
}