// RNA Web Service Reference
using RNA_Tools.Apex;

using System;
using System.Windows.Forms;

namespace RNA_Tools
{
    public partial class Form5 : Form
    {
        // Stores the type of import (Orders / Locations)
        private int Operation;

        private long LoginKey; // User EntityKey
        private long BUKey; // BU EntityKey
        private long RegionKey; // Region EntityKey
        private long SessionKey; // Session EntityKey
        private string RouteSessionDate; //Route Session Start Date

        // Array of Order Classes
        private OC[] Order_Class_Cache;

        // Array of Service Windows Types
        private TW[] Time_Window_Types_Cache;

        // Array of Service Windows Types
        private ST[] Service_Time_Types_Cache;

        // Stores Customer Locations List (Orders Import Required)
        private ListView Locations_Cache;

        // Order Class class
        private class OC
        {
            public long EntityKey;
            public string Id;
        }

        // Service Window Type cLass
        private class TW
        {
            public long EntityKey;
            public string Id;
        }

        // Service Time Type class
        private class ST
        {
            public long EntityKey;
            public string Id;
        }

        public Form5(string Title, ListView ToImport, int source, long Login_Key, long BU_Key, long Region_Key, long Session_Key,
            string Route_Session_Date, ListView Locs)
        {
            InitializeComponent();

            // Change the form title
            this.Text = Title;

            // Caches the source parameter needed to select between Orders and Locations
            this.Operation = source;

            // Caches the Login EntityKey
            this.LoginKey = Login_Key;

            // Caches the BU EntityKey
            this.BUKey = BU_Key;

            // Caches the Region EntityKey
            this.RegionKey = Region_Key;

            // Caches the Region EntityKey
            this.SessionKey = Session_Key;

            // Caches the Route Session Start Date
            this.RouteSessionDate = Route_Session_Date;

            // Caches the Locations List
            this.Locations_Cache = Locs;

            // Loads the items to be imported in the preview listview
            Load_List(ToImport);

            // Resize the columns of the listview
            AUX.Services.autoResizeColumns(listView1, 20);
        }

        // Loads all items to be import in the preview listview
        private void Load_List(ListView l)
        {
            listView1.Columns.Add("#");

            // Transfer columns
            foreach (ColumnHeader header in l.Columns)
            {
                listView1.Columns.Add((ColumnHeader)header.Clone());
            }

            int count = 0;

            // Transfer items
            foreach (ListViewItem item in l.Items)
            {
                // Create a new listviewitem to store the data
                ListViewItem i = new ListViewItem(); //

                // Adds the counter subitem to the lineitem
                count++;
                i.SubItems.Add("");
                i.SubItems[i.SubItems.Count - 2].Text = count.ToString();

                // Loads subitems into the listitem removing the first subitem - Worksheet Address
                for (int c = 1; c < item.SubItems.Count; c++)
                {
                    i.SubItems.Add(""); // Creates the subitem
                    i.SubItems[c].Text = item.SubItems[c].Text; // Fill the data into the bew subitem
                }

                // Add the new item to the list
                listView1.Items.Add(i);
            }
        }

        // Upload the records to the Web Services
        private void button1_Click(object sender, EventArgs e)
        {
            // Verifies if there are items to be imported
            if (listView1.Items.Count > 0)
            {
                // Initialize the variables
                progressBar1.Value = 0;
                progressBar1.Maximum = listView1.Items.Count;
                button1.Text = "Working...";

                // Clear Log
                listBox1.Items.Clear();

                // 1 -> Import orders
                if (Operation == 1)
                {
                    // Add information to the Log
                    listBox1.Items.Add("Importing Orders...");
                    listBox1.Refresh();

                    // Caches order classes to speed up the import process
                    if (Load_Order_Classes())
                    {
                        // Required fileds
                        int Idx_Identifier = ColumnId(listView1, "ID");
                        int Idx_Location = ColumnId(listView1, "LOCATION");
                        int Idx_Size1 = ColumnId(listView1, "SIZE1");
                        int Idx_Size2 = ColumnId(listView1, "SIZE2");
                        int Idx_Size3 = ColumnId(listView1, "SIZE3");
                        int Idx_Type = ColumnId(listView1, "TYPE");
                        int Idx_Class = ColumnId(listView1, "CLASS");
                        int Idx_Begin = ColumnId(listView1, "BEGIN");
                        int Idx_End = ColumnId(listView1, "END");

                        // Test if all required columns were found
                        if ((Idx_Identifier != -1) & (Idx_Location != -1) & (Idx_Size1 != -1) & (Idx_Size2 != -1) &
                            (Idx_Size3 != -1) & (Idx_Type != -1) & (Idx_Class != -1) & (Idx_Begin != -1) & (Idx_End != -1))
                        {
                            // Process all Orders in the listview
                            foreach (ListViewItem item in listView1.Items)
                            {
                                // Create the Order object
                                OrderSpec order = new OrderSpec();

                                // Check if the Order Type is Delivery or Pickup
                                if (item.SubItems[Idx_Type].Text == "DELIVERY")
                                {
                                    // Create the Task object
                                    DeliveryTaskSpec Delivery_Task = new DeliveryTaskSpec();

                                    // Retrive the EntityKey of the Service Location cache
                                    Delivery_Task.ServiceLocationEntityKey = Location_EntityKey(item.SubItems[Idx_Location].Text);

                                    // Initialize the quantities object
                                    Delivery_Task.Quantities = new Quantities()
                                    {
                                        Size1 = Convert.ToDouble(item.SubItems[Idx_Size1].Text),
                                        Size2 = Convert.ToDouble(item.SubItems[Idx_Size2].Text),
                                        Size3 = Convert.ToDouble(item.SubItems[Idx_Size3].Text),
                                    };

                                    //Create a Delivery Order
                                    order.TaskSpec = Delivery_Task;
                                }
                                else
                                {
                                    // Create the Task object
                                    PickupTaskSpec Pickup_Task = new PickupTaskSpec();

                                    // Retrive the EntityKey of the Service Location cache
                                    Pickup_Task.ServiceLocationEntityKey = Location_EntityKey(item.SubItems[Idx_Location].Text);

                                    // Initialize the quantities object
                                    Pickup_Task.Quantities = new Quantities()
                                    {
                                        Size1 = Convert.ToDouble(item.SubItems[Idx_Size1].Text),
                                        Size2 = Convert.ToDouble(item.SubItems[Idx_Size2].Text),
                                        Size3 = Convert.ToDouble(item.SubItems[Idx_Size3].Text),
                                    };

                                    // Create a Pickup Order
                                    order.TaskSpec = Pickup_Task;
                                }

                                // Mandatory parameters
                                order.ManagedByUserEntityKey = BUKey;
                                order.RegionEntityKey = RegionKey;
                                order.SessionEntityKey = SessionKey;

                                // Retrives the EntityKey of the Order Class
                                order.OrderClassEntityKey = OC_EntityKey(item.SubItems[Idx_Class].Text);

                                // Retrives the Order Id
                                order.Identifier = item.SubItems[Idx_Identifier].Text.Trim();

                                // Checks if the Begin date is empty and uses the Route Session Date
                                string beginDate = item.SubItems[Idx_Begin].Text.Trim();
                                if (beginDate.Length > 0)
                                {
                                    order.BeginDate = beginDate;
                                }
                                else
                                {
                                    order.BeginDate = RouteSessionDate;
                                }

                                // Checks if the End date is empty and uses the Route Session Date
                                string endDate = item.SubItems[Idx_End].Text.Trim();
                                if (beginDate.Length > 0)
                                {
                                    order.EndDate = endDate;
                                }
                                else
                                {
                                    order.EndDate = RouteSessionDate;
                                }

                                try
                                {
                                    // Call the Web Services to upload the new order
                                    long entity_key = RNA.WebServices.SaveOrder(order);

                                    listBox1.Items.Add("Order #" + order.Identifier + " - Added [EntityKey: " +
                                        entity_key.ToString() + "]");
                                }
                                catch (Exception Ex)
                                {
                                    listBox1.Items.Add("Order #" + order.Identifier + " - Warning: " + Ex.Message);
                                }

                                // Move Log to the last position
                                listBox1.SelectedIndex = listBox1.Items.Count - 1;

                                // Increments the progress bar
                                progressBar1.Value++;

                                // Refresh Form
                                this.Refresh();
                            }

                            listBox1.Items.Add("Finished.");
                        }
                        else
                        {
                            MessageBox.Show("One or more required columns was not found.", "Error");
                        }
                    }
                    else
                    {
                        MessageBox.Show("No Order Class found.", "Error");
                    }
                }
                else
                {
                    // Caches Service Window Types and Service Time Types to speed up the import process
                    if (Load_Service_Windows_Types() && Load_Service_Time_Types())
                    {
                        // Add information to the Log
                        listBox1.Items.Add("Importing Locations...");
                        listBox1.Refresh();

                        // Required fileds
                        int Idx_Identifier = ColumnId(listView1, "ID");
                        int Idx_Description = ColumnId(listView1, "DESCRIPTION");
                        int Idx_Address = ColumnId(listView1, "ADDRESS");
                        int Idx_City = ColumnId(listView1, "CITY");
                        int Idx_State = ColumnId(listView1, "STATE");
                        int Idx_Zip = ColumnId(listView1, "ZIP");
                        int Idx_Country = ColumnId(listView1, "COUNTRY");
                        int Idx_SW = ColumnId(listView1, "SERVICE_WINDOW");
                        int Idx_ST = ColumnId(listView1, "SERVICE_TIME");

                        // Test if all required columns were found
                        if ((Idx_Identifier != -1) & (Idx_Description != -1) & (Idx_Address != -1) &
                            (Idx_City != -1) & (Idx_State != -1) & (Idx_Zip != -1) & (Idx_Country != -1) &
                            (Idx_SW != -1) & (Idx_ST != -1))
                        {
                            //Create the Service Location object
                            ServiceLocation SrvLoc = new ServiceLocation();

                            //Create an Address object
                            Address LocAddress = new Address();

                            //Process all Locations in the listview
                            foreach (ListViewItem item in listView1.Items)
                            {
                                if (Location_EntityKey(item.SubItems[Idx_Identifier].Text) > 0)
                                {
                                    //Select Update Mode
                                    SrvLoc.Action = ActionType.Update;
                                }
                                else
                                {
                                    //Select Insert Mode
                                    SrvLoc.Action = ActionType.Add;
                                }

                                //Mandatory parameters
                                SrvLoc.BusinessUnitEntityKey = BUKey;
                                SrvLoc.RegionEntityKeys = new long[] { RegionKey };
                                SrvLoc.Identifier = item.SubItems[Idx_Identifier].Text;
                                SrvLoc.Description = item.SubItems[Idx_Description].Text;
                                SrvLoc.StandingDeliveryQuantities = new Quantities { };
                                SrvLoc.StandingPickupQuantities = new Quantities { };

                                //Build Location Address
                                LocAddress.AddressLine1 = item.SubItems[Idx_Address].Text; //Address 1
                                LocAddress.Locality = new Locality
                                {
                                    AdminDivision2 = item.SubItems[Idx_City].Text, //City
                                    AdminDivision1 = item.SubItems[Idx_State].Text, //State
                                    PostalCode = item.SubItems[Idx_Zip].Text, //Postal Code
                                    CountryISO3Abbr = item.SubItems[Idx_Country].Text // Country Code
                                };

                                // Load Address into Location object
                                SrvLoc.Address = LocAddress;

                                // Retrives the EntityKey of the Service Window Type
                                SrvLoc.TimeWindowTypeEntityKey = TW_EntityKey(item.SubItems[Idx_SW].Text);
                                // Retrives the EntityKey of the Service Time Type
                                SrvLoc.ServiceTimeTypeEntityKey = ST_EntityKey(item.SubItems[Idx_ST].Text);

                                try
                                {
                                    // Call the Web Services to upload the new Location
                                    long entity_key = RNA.WebServices.SaveLocation(SrvLoc);

                                    listBox1.Items.Add("Location: " + SrvLoc.Identifier + " - Added [EntityKey: " +
                                        entity_key.ToString() + "]");
                                }
                                catch (Exception Ex)
                                {
                                    listBox1.Items.Add("Location: " + SrvLoc.Identifier + " - Warning: " + Ex.Message);
                                }

                                // Move Log to the last position
                                listBox1.SelectedIndex = listBox1.Items.Count - 1;

                                // Increments the progress bar
                                progressBar1.Value++;

                                // Refresh Form
                                this.Refresh();
                            }
                        }
                        else
                        {
                            MessageBox.Show("One or more required columns was not found.", "Error");
                        }
                    }
                    else
                    {
                        MessageBox.Show("No Service Windows Type or Service Time Type found.", "Error");
                    }

                }

                button1.Text = "Run";
            }
            else
            {
                MessageBox.Show("No items to import.", "Error");
            }          
        }

        // Returns the Id from the column
        private int ColumnId(ListView l, string Name)
        {
            // Search the index of the column name
            foreach (ColumnHeader c in l.Columns)
            {
                // Check the column name
                if (c.Text.Trim().ToUpper() == Name.Trim().ToUpper())
                {
                    return c.Index;
                }
            }

            // Column name not found
            return -1;
        }

        // Returns the EntityKey from the Location Id
        private long Location_EntityKey(string Id)
        {
            long r = -1;

            //Search EntityKey in the Locations ListView
            foreach (ListViewItem item in Locations_Cache.Items)
            {
                // Check the Customer Id
                if (item.SubItems[1].Text == Id)
                {
                    // Gets the EntityKey from the Customer Id
                    r = Convert.ToInt64(item.SubItems[0].Text);
                    break;
                }
            }

            return r;
        }

        // Returns the EntityKey from the Order Class Id
        private long OC_EntityKey(string Id)
        {
            long r = -1;

            // Search the index of the column name
            foreach (OC oc in Order_Class_Cache)
            {
                // Check the column name
                if (oc.Id.Trim().ToUpper() == Id.Trim().ToUpper())
                {
                    r = oc.EntityKey;
                    break;
                }
            }

            return r;
        }

        // Returns the EntityKey from the Service Windows Type Id
        private long TW_EntityKey(string Id)
        {
            long r = -1;

            // Search the index of the column name
            foreach (TW tw in Time_Window_Types_Cache)
            {
                // Check the column name
                if (tw.Id.Trim().ToUpper() == Id.Trim().ToUpper())
                {
                    r = tw.EntityKey;
                    break;
                }
            }

            return r;
        }

        // Returns the EntityKey from the Service Time Type Id
        private long ST_EntityKey(string Id)
        {
            long r = -1;

            // Search the index of the column name
            foreach (ST st in Service_Time_Types_Cache)
            {
                // Check the column name
                if (st.Id.Trim().ToUpper() == Id.Trim().ToUpper())
                {
                    r = st.EntityKey;
                    break;
                }
            }

            return r;
        }

        // Caches the Order Classes of the region
        private bool Load_Order_Classes()
        {
            try
            {
                // Loads the Order Classes data
                OrderClass[] OC_Result = RNA.WebServices.RetrieveOrderClasses();

                // Get the number of elements retrived
                int n = OC_Result.Length;

                // Checks there are elements to process
                if (n > 0)
                {
                    // Redefines the array with the number of objects found in the Web Services
                    Order_Class_Cache = new OC[n];

                    // Resets the N variable to be reused
                    n = 0;

                    // Loads the Order Classes data into the Order Classes Array
                    foreach (OrderClass oc in OC_Result)
                    {
                        Order_Class_Cache[n] = new OC();
                        Order_Class_Cache[n].EntityKey = oc.EntityKey;
                        Order_Class_Cache[n].Id = oc.Identifier;
                        n++;
                    }

                    return true;
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Error");
            }

            return false;
        }

        // Caches the Service Windows of the region
        private bool Load_Service_Windows_Types()
        {
            try
            {
                // Loads the Service Window Types data
                TimeWindowType[] TW_Result = RNA.WebServices.RetreiveTimeWindowTypes();

                // Get the number of elements retrived
                int n = TW_Result.Length;

                // Checks there are elements to process
                if (n > 0)
                {
                    // Redefines the array with the number of objects found in the Web Services
                    Time_Window_Types_Cache = new TW[n];

                    // Resets the N variable to be reused
                    n = 0;

                    // Loads the Service Window Types data into the Service Window Types array
                    foreach (TimeWindowType tw in TW_Result)
                    {
                        Time_Window_Types_Cache[n] = new TW();
                        Time_Window_Types_Cache[n].EntityKey = tw.EntityKey;
                        Time_Window_Types_Cache[n].Id = tw.Identifier;
                        n++;
                    }

                    return true;
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Error");
            }

            return false;
        }

        // Caches the Service Time Types of the region
        private bool Load_Service_Time_Types()
        {
            try
            {
                // Loads the Service Time Types data
                ServiceTimeType[] ST_Result = RNA.WebServices.RetreiveServiceTimeTypes();

                // Get the number of elements retrived
                int n = ST_Result.Length;

                // Checks there are elements to process
                if (n > 0)
                {
                    // Redefines the array with the number of objects found in the Web Services
                    Service_Time_Types_Cache = new ST[n];

                    // Resets the N variable to be reused
                    n = 0;

                    // Loads the Service Time Types data into the Service Time Types array
                    foreach (ServiceTimeType st in ST_Result)
                    {
                        Service_Time_Types_Cache[n] = new ST();
                        Service_Time_Types_Cache[n].EntityKey = st.EntityKey;
                        Service_Time_Types_Cache[n].Id = st.Identifier;
                        n++;
                    }

                    return true;
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Error");
            }

            return false;
        }
    }
}
