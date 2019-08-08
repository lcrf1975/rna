// RNA Web Service Reference
using RNA_Tools.Apex;

using System;
using System.Configuration;
using System.Linq;

namespace RNA
{
    class WebServices
    {
        public static SessionHeader SessionHeader; // Stores the Webservie session header
        public static QueryServiceClient QueryServiceClient; // Object to handle the Query Client Service from the Webservices
        public static SingleRegionContext RegionContext; // Stores the Region context
        public static MappingServiceClient MappingServiceClient; // Object to handle the Mapping Client Service from the Webservices
        public static RoutingServiceClient RoutingServiceClient; // Object to handle the Routing Client Service from the Webservices

        // Login to RNA Webservice - Needs a Webservice valid user
        public static string[] Login()
        {
            try
            {
                // Build Application Information. The ClientApplicationIdentifier GUID is the application ID 
                // for integration users and should not be changed
                ClientApplicationInfo AppInfo = new ClientApplicationInfo
                {
                    ClientApplicationIdentifier = new Guid(ConfigurationManager.AppSettings["ClientApplicationIdentifier"])
                };

                // Call the Web Services
                LoginServiceClient loginServiceClient = new LoginServiceClient();
                LoginResult loginResult = loginServiceClient.Login(
                    ConfigurationManager.AppSettings["Login"],
                    ConfigurationManager.AppSettings["Password"],
                    new CultureOptions(),
                    AppInfo);

                // Checks if the process returned a valid result
                if (loginResult == null)
                {
                    throw new Exception("Login failed.");
                }
                else
                {
                    // Stores the Webservice session header
                    SessionHeader = new SessionHeader { SessionGuid = loginResult.UserSession.Guid };

                    // Enable the Query Client Service
                    QueryServiceClient = new QueryServiceClient("BasicHttpBinding_IQueryService", loginResult.QueryServiceUrl);
                    string[] r = new string[] { loginResult.User.EntityKey.ToString(), loginResult.User.EmailAddress.ToString() };

                    return r;
                }

            }
            catch (Exception Ex)
            {
                throw new Exception(Ex.Message);
            }
        }

        // Retrives all BU from the connected users
        public static BusinessUnit[] RetrieveBusinessUnits()
        {
            try
            {
                // Call the Web Services
                RetrievalResults retrievalResults = QueryServiceClient.RetrieveBusinessUnitsGrantingPermissions(
                    SessionHeader,
                    new RolePermission[] { },
                    false);

                // Checks if the process returned a valid result
                if (retrievalResults.Items == null)
                {
                    throw new Exception("Retrieve Business Units failed.");
                }
                else if (retrievalResults.Items.Length == 0)
                {
                    throw new Exception("No Business Units found.");
                }
                else
                {
                    // Retrieves Business Units completed successfully
                    return retrievalResults.Items.Cast<BusinessUnit>().ToArray();
                }
            }
            catch (Exception Ex)
            {
                throw new Exception(Ex.Message);
            }
        }

        // Retrives all Regions from the current BU 
        public static Region[] RetrieveRegions()
        {
            try
            {
                // Call the Web Services
                RetrievalResults retrievalResults = QueryServiceClient.RetrieveRegionsGrantingPermissions(
                    SessionHeader,
                    new RolePermission[] { },
                    false);

                // Checks if the process returned a valid result
                if (retrievalResults.Items == null)
                {
                    throw new Exception("Retrieve Regions failed.");
                }
                else if (retrievalResults.Items.Length == 0)
                {
                    throw new Exception("No Regions found.");
                }
                else
                {
                    // Retrieves Regions completed successfully
                    return retrievalResults.Items.Cast<Region>().ToArray();
                }
            }
            catch (Exception Ex)
            {
                throw new Exception(Ex.Message);
            }
        }

        // Retrieves and creates the URL Webservices objects
        public static string[] RetrieveUrls(long buId, long regionId)
        {
            try
            {
                // Object with the search parameters
                RegionContext = new SingleRegionContext
                {
                    BusinessUnitEntityKey = buId,
                    RegionEntityKey = regionId
                };

                // Call the Web Services
                UrlSet urlSet = QueryServiceClient.RetrieveUrlsForContext(
                    SessionHeader,
                    RegionContext);

                // Checks if the returned data is valid
                if (urlSet == null)
                {
                    throw new Exception("Retrieve Urls failed.");
                }
                else
                {
                    // Creates URL Webservice objects
                    MappingServiceClient = new MappingServiceClient("BasicHttpBinding_IMappingService", urlSet.MappingService);
                    RoutingServiceClient = new RoutingServiceClient("BasicHttpBinding_IRoutingService", urlSet.RoutingService);

                    // Stores the URL address
                    string[] r = new string[] { urlSet.MappingService, urlSet.RoutingService };

                    // Retrieves URL completed successfully
                    return r;
                }
            }
            catch (Exception Ex)
            {
                throw new Exception(Ex.Message);
            }
        }

        // Returns all Routing Sessions from the selected Region
        public static DailyRoutingSession[] RetrieveRouteSessions()
        {
            try
            {
                // Call the Web Services
                RetrievalResults retrievalResults = QueryServiceClient.Retrieve(
                    SessionHeader,
                    RegionContext,
                    new RetrievalOptions
                    {
                        PropertyInclusionMode = PropertyInclusionMode.All,
                        Type = "DailyRoutingSession"
                    });

                // Checks if the process returned a valid result
                if (retrievalResults.Items == null)
                {
                    throw new Exception("Retrieve Route Sessions failed.");
                }
                else if (retrievalResults.Items.Length == 0)
                {
                    throw new Exception("No Route Sessions found for the select user.");
                }
                else
                {
                    // Retrieves Route Sessions completed successfully
                    return retrievalResults.Items.Cast<DailyRoutingSession>().ToArray();
                }
            }
            catch (Exception Ex)
            {
                throw new Exception(Ex.Message);
            }
        }

        // Returns all Routes from the selected Route Session
        public static Route[] RetrieveRoutes(long routingSessionEntityKey)
        {
            try
            {
                // Call the Web Services
                RetrievalResults retrievalResults = QueryServiceClient.Retrieve(
                    SessionHeader,
                    RegionContext,
                    new RetrievalOptions
                    {
                        Expression = new AndExpression
                        {
                            Expressions = new SimpleExpressionBase[]
                            {
                                new EqualToExpression
                                {
                                    Left = new PropertyExpression { Name = "RoutingSessionEntityKey" },
                                    Right = new ValueExpression { Value = routingSessionEntityKey }
                                }
                            }
                        },
                        PropertyInclusionMode = PropertyInclusionMode.All,
                        Type = "Route"
                    });

                // Checks if the process returned a valid result
                if (retrievalResults.Items == null)
                {
                    throw new Exception("Retrieve Routes failed for the Route Session " + routingSessionEntityKey.ToString() + ".");
                }
                else if (retrievalResults.Items.Length == 0)
                {
                    throw new Exception("No Routes found in the Route Session " + routingSessionEntityKey.ToString() + ".");
                }
                else
                {
                    // Retrieves Routes completed successfully
                    return retrievalResults.Items.Cast<Route>().ToArray();
                }
            }
            catch (Exception Ex)
            {
                throw new Exception(Ex.Message);
            }
        }

        // Returns all Service Locations from the select region
        public static ServiceLocation[] RetrieveServiceLocation()
        {
            try
            {
                // Call the Web Services
                RetrievalResults retrievalResults = QueryServiceClient.Retrieve(
                    SessionHeader,
                    RegionContext,
                    new RetrievalOptions
                    {
                        PropertyInclusionMode = PropertyInclusionMode.All,
                        Type = "ServiceLocation"
                    });

                // Checks if the process returned a valid result
                if (retrievalResults.Items == null)
                {
                    throw new Exception("Retrieve Service Location failed.");
                }
                else if (retrievalResults.Items.Length == 0)
                {
                    throw new Exception("No Service Location found for selected Region.");
                }
                else
                {
                    // Retrieves Service Location completed successfully
                    return retrievalResults.Items.Cast<ServiceLocation>().ToArray();
                }
            }
            catch (Exception Ex)
            {
                throw new Exception(Ex.Message);
            }
        }

        // Returns all Order Classes from the select region
        public static OrderClass[] RetrieveOrderClasses()
        {
            try
            {
                // Call the Web Services
                RetrievalResults retrievalResults = QueryServiceClient.Retrieve(
                    SessionHeader,
                    RegionContext,
                    new RetrievalOptions
                    {
                        PropertyInclusionMode = PropertyInclusionMode.All,
                        Type = "OrderClass"
                    });

                // Checks if the process returned a valid result
                if (retrievalResults.Items == null)
                {
                    throw new Exception("Retrieve Order Class failed.");
                }
                else if (retrievalResults.Items.Length == 0)
                {
                    throw new Exception("No Ordewr Class found for selected Region.");
                }
                else
                {
                    // Retrieves Service Location completed successfully
                    return retrievalResults.Items.Cast<OrderClass>().ToArray();
                }
            }
            catch (Exception Ex)
            {
                throw new Exception(Ex.Message);
            }
        }

        // Returns all Service Window Types from the select region
        public static TimeWindowType[] RetreiveTimeWindowTypes()
        {
            try
            {
                // Call the Web Services
                RetrievalResults retrievalResults = QueryServiceClient.Retrieve(
                    SessionHeader,
                    RegionContext,
                    new RetrievalOptions
                    {
                        PropertyInclusionMode = PropertyInclusionMode.All,
                        Type = "TimeWindowType"
                    });

                // Checks if the process returned a valid result
                if (retrievalResults.Items == null)
                {
                    throw new Exception("Retrieve Service Windows Types failed.");
                }
                else if (retrievalResults.Items.Length == 0)
                {
                    throw new Exception("No Service Widnows Type found for selected Region.");
                }
                else
                {
                    // Retrieves Service Location completed successfully
                    return retrievalResults.Items.Cast<TimeWindowType>().ToArray();
                }
            }
            catch (Exception Ex)
            {
                throw new Exception(Ex.Message);
            }
        }

        // Returns all Service Time Types from the select region
        public static ServiceTimeType[] RetreiveServiceTimeTypes()
        {
            try
            {
                // Call the Web Services
                RetrievalResults retrievalResults = QueryServiceClient.Retrieve(
                    SessionHeader,
                    RegionContext,
                    new RetrievalOptions
                    {
                        PropertyInclusionMode = PropertyInclusionMode.All,
                        Type = "ServiceTimeType"
                    });

                // Checks if the process returned a valid result
                if (retrievalResults.Items == null)
                {
                    throw new Exception("Retrieve Service Windows Types failed.");
                }
                else if (retrievalResults.Items.Length == 0)
                {
                    throw new Exception("No Service Widnows Type found for selected Region.");
                }
                else
                {
                    // Retrieves Service Location completed successfully
                    return retrievalResults.Items.Cast<ServiceTimeType>().ToArray();
                }
            }
            catch (Exception Ex)
            {
                throw new Exception(Ex.Message);
            }
        }

        // Returns all Service Order from the selected Route Session 
        public static Order[] RetrieveServiceOrders(long sessionEntityKey)
        {
            try
            {
                // Call the Web Services
                RetrievalResults retrievalResults = QueryServiceClient.Retrieve(
                    SessionHeader,
                    RegionContext,
                    new RetrievalOptions
                    {
                        Expression = new AndExpression
                        {
                            Expressions = new SimpleExpressionBase[]
                            {
                                new EqualToExpression
                                {
                                    Left = new PropertyExpression { Name = "SessionEntityKey" },
                                    Right = new ValueExpression { Value = sessionEntityKey }
                                }
                            }
                        },
                        PropertyInclusionMode = PropertyInclusionMode.All,
                        Type = "Order"
                    });

                // Checks if the process returned a valid result
                if (retrievalResults.Items == null)
                {
                    throw new Exception("Retrieve Service Orders failed.");
                }
                else if (retrievalResults.Items.Length == 0)
                {
                    throw new Exception("No Service Orders found for the Route Session "
                        + sessionEntityKey.ToString() + ".");
                }
                else
                {
                    // Retrieves Service Orders completed successfully
                    return retrievalResults.Items.Cast<Order>().ToArray();
                }
            }
            catch (Exception Ex)
            {
                throw new Exception(Ex.Message);
            }
        }

        // Searchs the Depot Name
        public static string[] SearchDepot(long depotEntityKey)
        {
            string[] r = new string[4];

            // Means: No depot found.
            r[0] = "NOT FOUND"; // Name
            r[1] = "0.0000"; // Lat
            r[2] = "0.0000"; // Lon
            r[3] = "DEPOT ADDRESS NOT FOUND"; // Address
            // End

            try
            {
                // Call the Web Services
                RetrievalResults retrievalResults = QueryServiceClient.Retrieve(
                    SessionHeader,
                    RegionContext,
                     new RetrievalOptions
                     {
                         PropertyInclusionMode = PropertyInclusionMode.All,
                         Type = "Depot"
                     });

                // Checks if the process returned a valid result
                if (retrievalResults.Items != null | retrievalResults.Items.Length > 0)
                {
                    // Loads all Depots into an array
                    Depot[] Depots = retrievalResults.Items.Cast<Depot>().ToArray();

                    // Finds the Depot by EntityKey
                    foreach (Depot depot in Depots)
                    {
                        if (depot.EntityKey == depotEntityKey) // Depot Found
                        {
                            r[0] = depot.Identifier;
                            r[1] = depot.Address.AddressLine1 + ", "
                                + depot.Address.Locality.AdminDivision3 + ", " // City
                                + depot.Address.Locality.AdminDivision1 + ", " // State
                                + depot.Address.Locality.PostalCode;
                            r[2] = (depot.Coordinate.Latitude * 0.000001).ToString();
                            r[3] = (depot.Coordinate.Longitude * 0.000001).ToString();

                            // Exit loop
                            break;
                        }
                    }
                }

                return r;
            }
            catch (Exception Ex)
            {
                throw new Exception(Ex.Message);
            }
        }

        // Geocodes an Address
        public static string[] Geocode(Address address)
        {
            try
            {
                // Call the Web Services
                GeocodeResult[] geocodeResults = MappingServiceClient.Geocode(
                    SessionHeader,
                    RegionContext,
                    new GeocodeCriteria
                    {
                        NamedPlaces = new NamedPlace[] { new NamedPlace { PlaceAddress = address } },
                    },
                    new GeocodeOptions
                    {
                        NetworkArcCandidatePropertyInclusionMode = PropertyInclusionMode.All,
                        NetworkPOICandidatePropertyInclusionMode = PropertyInclusionMode.All,
                        NetworkPointAddressCandidatePropertyInclusionMode = PropertyInclusionMode.All
                    });

                int c = 0;
                string[] r;

                // Checks if the process returned a valid result
                if (geocodeResults[0].Results == null || geocodeResults[0].Results.Length == 0)
                {
                    // No data found
                    r = new string[1];
                    r[c] = "No results found|||";
                }
                else
                {
                    // Returns an array with all Geocode data found
                    r = new string[geocodeResults[0].Results.Length];
                    foreach (GeocodeCandidate geocodeResult in geocodeResults[0].Results)
                    {
                        double lat = geocodeResult.Coordinate.Latitude * 0.000001;
                        double lon = geocodeResult.Coordinate.Longitude * 0.000001;
                        r[c] = lat.ToString() + "|"
                            + lon.ToString() + "|"
                            + geocodeResult.Score.ToString() + "|"
                            + geocodeResult.GeocodeAccuracy_Quality.ToString();
                        c++;
                    }
                }

                return r;
            }
            catch (Exception Ex)
            {
                throw new Exception(Ex.Message);
            }
        }

        // Generates the Travel Path
        public static TravelPathResult TravelPath(Coordinate[] Coord)
        {
            try
            {
                // Call the Web Services
                TravelPathResult pathResult = MappingServiceClient.GenerateTravelPath(
                    SessionHeader,
                    RegionContext,
                    new GenerateTravelPathOptions
                    {
                        MeasurementOptions = new MeasurementOptions
                        {
                            DistanceUnit_DistanceUnit = DistanceUnit.Kilometers.ToString(),
                            FuelEconomyUnit_FuelEconomyUnit = FuelEconomyUnit.LitersPerHundredKilometers.ToString(),
                            LengthUnit_LengthUnit = LengthUnit.Centimeters.ToString(),
                            LiquidVolumeUnit_LiquidVolumeUnit = LiquidVolumeUnit.Liters.ToString(),
                            SpeedUnit_SpeedUnit = SpeedUnit.KilometersPerHour.ToString(),
                            WeightUnit_WeightUnit = WeightUnit.Kilograms.ToString()
                        },
                        PropertyOptions = new TravelPathResultPropertyOptions
                        {
                            ArcIDs = false,
                            DestinationPathErrors = false,
                            DestinationPointIndices = false,
                            Directions = false,
                            Path = true     
                        },
                        CalculatorOptions = new TravelPathCalculatorOptions
                        {
                            LoadEndpointsToNodes = false,
                            LoadEndpointsToLoadableRoads = false
                        },
                        Coordinates = Coord
                    });

                return pathResult; // Returns the Travel Path result
            }
            catch (Exception Ex)
            {
                throw new Exception(Ex.Message);
            }
        }

        // Creates a new Route Session
        public static long NewRoutingSession(DailyRoutingSession dailyRoutingSession)
        {
            try
            {
                // Call the Web Services
                SaveResult[] saveResults = RoutingServiceClient.Save(
                    SessionHeader, RegionContext,
                    new AggregateRootEntity[] { dailyRoutingSession },
                    new SaveOptions
                    {
                        InclusionMode = PropertyInclusionMode.All
                    });

                if (saveResults[0].Error != null) // Error creating the New Route Session
                {
                    throw new Exception(saveResults[0].Error.Code.ErrorCode_Status);
                }

                return saveResults[0].EntityKey; // New Route Session created

            }
            catch (Exception Ex)
            {
                throw new Exception(Ex.Message);
            }

        }

        // Creates a new Order
        public static long SaveOrder(OrderSpec orderSpec)
        {
            try
            {
                // Call the Web Services
                SaveResult[] saveResults = RoutingServiceClient.SaveOrders(
                    SessionHeader, RegionContext,
                    new OrderSpec[] { orderSpec },
                    new SaveOptions
                    {
                        InclusionMode = PropertyInclusionMode.All
                    });

                if (saveResults[0].Error != null) // Error creating the Order
                {
                    throw new Exception(saveResults[0].Error.Code.ErrorCode_Status);
                }

                return saveResults[0].EntityKey; // New Order created

            }
            catch (Exception Ex)
            {
                throw new Exception(Ex.Message);
            }
        }

        public static long SaveLocation(ServiceLocation serviceLocation)
        {
            try
            {
                // Call the Web Services
                SaveResult[] saveResults = RoutingServiceClient.Save(
                    SessionHeader,
                    RegionContext,
                    new AggregateRootEntity[] { serviceLocation },
                    new SaveOptions
                    {
                        InclusionMode = PropertyInclusionMode.All
                    });

                if (saveResults[0].Error != null) // Error creating the Location
                {
                    throw new Exception(saveResults[0].Error.Code.ErrorCode_Status);
                }

                return saveResults[0].EntityKey; // New Location created
            }
            catch (Exception Ex)
            {
                throw new Exception(Ex.Message);
            }
        }
    }
}