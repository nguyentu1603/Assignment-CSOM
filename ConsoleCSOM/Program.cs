using Microsoft.Extensions.Configuration;
using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;
using System.Linq;
using Microsoft.SharePoint.Client.Taxonomy;
using ConsoleCSOM.Helpers;
using System.Collections.Generic;
using ConsoleCSOM.Models;

namespace ConsoleCSOM
{
    class SharepointInfo
    {
        public string AdminSiteUrl { get; set; }
        public string SiteUrl { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
    }

    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                using (var clientContextHelper = new ClientContextHelper())
                {
                    ClientContext ctx = GetContext(clientContextHelper);
                    ctx.Load(ctx.Web);
                    await ctx.ExecuteQueryAsync();
                    //await CreateList(ctx);
                    //await CreateTermGroup(ctx);
                    //await CreateTermSet(ctx);
                    //await CreateTerm(ctx);
                    //await CreateField(ctx);
                    //await CreateContentTypes(ctx);
                    //await AddContentTypeToList(ctx);
                    //await AddItemToList(ctx);
                    //await UpdateAboutField(ctx);
                    //await UpdateCityField(ctx);
                    //await CAMLQueryAsync(ctx);
                    //await ListViewCSOMOrder(ctx);
                    //await UpdateList(ctx);
                    //await AddAuthorToSCOMList(ctx);
                    //await CreateCitiesField(ctx);
                    await AddListWithCitiesField(ctx);
                    Console.WriteLine($"Site {ctx.Web.Title}");
                }
                Console.WriteLine($"Press Any Key To Stop!");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        static ClientContext GetContext(ClientContextHelper clientContextHelper)
        {
            var builder = new ConfigurationBuilder().AddJsonFile($"appsettings.json", true, true);
            IConfiguration config = builder.Build();
            var info = config.GetSection("SharepointInfo").Get<SharepointInfo>();
            return clientContextHelper.GetContext(new Uri(info.SiteUrl), info.Username, info.Password);
        }

        private static async Task CreateList(ClientContext ctx)
        {
            ListCreationInformation creationInfo = new ListCreationInformation();
            creationInfo.Title = "CSOM Test";
            creationInfo.Description = "List for CSOM Test";
            creationInfo.TemplateType = (int)ListTemplateType.GenericList;
            List newList = ctx.Web.Lists.Add(creationInfo);
            ctx.Load(newList);
            // Execute the query to the server.
            await ctx.ExecuteQueryAsync();
        }

        private static async Task CreateTermGroup(ClientContext ctx)
        {
            string termGroupName = "CSOM-Test";
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            //Create Term Group
            TermGroup termGroup = termStore.CreateGroup(termGroupName, Guid.NewGuid());
            ctx.Load(termGroup);
            // Execute the query to the server.
            await ctx.ExecuteQueryAsync();
        }

        private static async Task CreateTermSet(ClientContext ctx)
        {
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            //Create Term Set
            TermGroup termGroup = termStore.Groups.GetByName("CSOM-Test");
            TermSet termSet = termGroup.CreateTermSet("city-NguyenAnhTu", Guid.NewGuid(), Constants.LCID_ENGLISH);
            // Execute the query to the server.
            await ctx.ExecuteQueryAsync();
        }
        private static async Task CreateTerm(ClientContext ctx)
        {
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            //Get Term Group & Term Set
            TermGroup termGroup = termStore.Groups.GetByName("CSOM-Test");
            TermSet termSet = termGroup.TermSets.GetByName("city-NguyenAnhTu");
            //Create new Term
            Term hcmTerm = termSet.CreateTerm("Ho Chi Minh", Constants.LCID_ENGLISH, Guid.NewGuid());
            Term stockHolmTerm = termSet.CreateTerm("Stockholm", Constants.LCID_ENGLISH, Guid.NewGuid());
            // Execute the query to the server.
            await ctx.ExecuteQueryAsync();
        }

        private static async Task CreateField(ClientContext ctx)
        {
            //Create About (Text) Field
            ctx.Site.RootWeb.Fields.AddFieldAsXml("<Field DisplayName='About' Name='About' Group='CSOM Test Group' Type='Text' />", false, AddFieldOptions.AddFieldInternalNameHint);
            //Create City (TaxonomyFieldType) Field
            ctx.Site.RootWeb.Fields.AddFieldAsXml("<Field DisplayName='City' Name='City' Group='CSOM Test Group' Type='TaxonomyFieldType' />", false, AddFieldOptions.AddFieldInternalNameHint);
            await ctx.ExecuteQueryAsync();
        }

        private static async Task CreateContentTypes(ClientContext ctx)
        {
            //Create Content Types
            ctx.Site.RootWeb.ContentTypes.Add(new ContentTypeCreationInformation
            {
                Name = "CSOM Test Content Type",
                Group = "CSOM Test Content Types"
            });
            await ctx.ExecuteQueryAsync();
            // Get all the content types from current site
            ContentTypeCollection contentTypeCollection = ctx.Web.ContentTypes;
            // Load content type collection
            ctx.Load(contentTypeCollection);
            await ctx.ExecuteQueryAsync();
            // Give content type name over here
            ContentType targetContentType = (from contentType in contentTypeCollection where contentType.Name == "CSOM Test Content Type" select contentType).FirstOrDefault();
            ctx.Load(targetContentType);
            await ctx.ExecuteQueryAsync();
            Field aboutField = ctx.Site.RootWeb.Fields.GetByInternalNameOrTitle("About");
            Field cityField = ctx.Site.RootWeb.Fields.GetByInternalNameOrTitle("City");
            targetContentType.FieldLinks.Add(new FieldLinkCreationInformation
            {
                Field = aboutField
            });
            targetContentType.FieldLinks.Add(new FieldLinkCreationInformation
            {
                Field = cityField,
            });
            targetContentType.Update(true);
            await ctx.ExecuteQueryAsync();
        }

        private static async Task AddContentTypeToList(ClientContext ctx)
        {
            // Get all the content types from current site
            ContentTypeCollection contentTypeCollection = ctx.Site.RootWeb.ContentTypes;
            ctx.Load(contentTypeCollection);
            await ctx.ExecuteQueryAsync();
            //Add Content Type To List
            ContentType targetContentType = (from contentType in contentTypeCollection where contentType.Name == "CSOM Test Content Type" select contentType).FirstOrDefault();
            List targetList = ctx.Web.Lists.GetByTitle("CSOM Test");
            targetList.ContentTypes.AddExistingContentType(targetContentType);
            targetList.Update();
            ctx.Web.Update();
            await ctx.ExecuteQueryAsync();
            //Load current Content Types in List
            var currentCtOrder = targetList.ContentTypes;
            ctx.Load(currentCtOrder, coll => coll.Include(
                                    ct => ct.Name,
                                    ct => ct.Id));
            await ctx.ExecuteQueryAsync();
            IList<ContentTypeId> reverseOrder = (from ct in currentCtOrder where ct.Name.Equals("CSOM Test Content Type", StringComparison.OrdinalIgnoreCase) select ct.Id).ToList();
            targetList.RootFolder.UniqueContentTypeOrder = reverseOrder;
            targetList.RootFolder.Update();
            targetList.Update();
            await ctx.ExecuteQueryAsync();
        }

        private static async Task AddItemToList(ClientContext ctx)
        {
            //Create Item List
            List<ItemList> csomList = new List<ItemList>();
            csomList.Add(new ItemList("Test 1", "About 1", "City 1"));
            csomList.Add(new ItemList("Test 2", "About 2", "City 2"));
            csomList.Add(new ItemList("Test 3", "About 3", "City 3"));
            csomList.Add(new ItemList("Test 4", "About 4", "City 4"));
            csomList.Add(new ItemList("Test 5", "About 5", "City 5"));
            List targetList = ctx.Web.Lists.GetByTitle("CSOM Test");
            // Get Target List
            foreach (var item in csomList)
            {
                ListItemCreationInformation iteminfo = new ListItemCreationInformation();
                ListItem newListItem = targetList.AddItem(iteminfo);
                newListItem["Title"] = item.Title;
                newListItem["About"] = item.About;
                // City Field need to implement
                newListItem.Update();
                await ctx.ExecuteQueryAsync();
            }
            await ctx.ExecuteQueryAsync();
        }

        private static async Task UpdateAboutField(ClientContext ctx)
        {
            Field aboutField = ctx.Site.RootWeb.Fields.GetByInternalNameOrTitle("About");
            aboutField.DefaultValue = "about default";
            aboutField.UpdateAndPushChanges(true);
            await ctx.ExecuteQueryAsync();
            AddTwoNewItems(ctx);
        }

        private static async Task UpdateCityField(ClientContext ctx)
        {
            Guid termStoreId = Guid.Empty;
            Guid termSetId = Guid.Empty;
            GetTaxonomyFieldInfo(ctx, out termStoreId, out termSetId);

            Field cityField = ctx.Site.RootWeb.Fields.GetByInternalNameOrTitle("City");

            // Retrieve as Taxonomy Field - Add Term Set to Taxonomy Field
            TaxonomyField taxonomyField = ctx.CastTo<TaxonomyField>(cityField);
            taxonomyField.SspId = termStoreId;
            taxonomyField.TermSetId = termSetId;
            taxonomyField.TargetTemplate = String.Empty;
            taxonomyField.AnchorId = Guid.Empty;
            taxonomyField.Update();
            await ctx.ExecuteQueryAsync();

            //Get All Terms
            // Get the TaxonomySession
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            // Get the term store by name
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            // Get the term group by Name
            TermGroup termGroup = termStore.Groups.GetByName("CSOM-Test");
            // Get the term set by Name
            TermSet termSet = termGroup.TermSets.GetByName("city-NguyenAnhTu");
            var term = termSet.Terms.GetByName("Ho Chi Minh");
            ctx.Load(term);
            await ctx.ExecuteQueryAsync();

            //Set Default Value
            TaxonomyFieldValue defaultValue = new TaxonomyFieldValue();
            defaultValue.WssId = -1;
            defaultValue.Label = term.Name;
            // GUID should be stored lowercase, otherwise it will not work in Office 2010
            defaultValue.TermGuid = term.Id.ToString().ToLower();

            // Get the Validated String for the taxonomy value
            var validatedValue = taxonomyField.GetValidatedString(defaultValue);
            await ctx.ExecuteQueryAsync();

            // Set the selected default value for the site column
            taxonomyField.DefaultValue = validatedValue.Value;
            taxonomyField.UserCreated = false;
            taxonomyField.UpdateAndPushChanges(true);
            await ctx.ExecuteQueryAsync();

            AddTwoNewItems(ctx);
        }

        private static void AddTwoNewItems(ClientContext ctx)
        {
            // Add 2 New Items
            List targetList = ctx.Web.Lists.GetByTitle("CSOM Test");
            ListItemCreationInformation iteminfo = new ListItemCreationInformation();
            ListItem newListItem = targetList.AddItem(iteminfo);
            newListItem["Title"] = "Test " + DateTime.Now.ToString();
            newListItem.Update();
            ListItem newListItem2 = targetList.AddItem(iteminfo);
            newListItem2["Title"] = "Test " + DateTime.Now.ToString();
            newListItem2.Update();
            ctx.ExecuteQuery();
        }

        private static void GetTaxonomyFieldInfo(ClientContext ctx, out Guid termStoreId, out Guid termSetId)
        {
            termStoreId = Guid.Empty;
            termSetId = Guid.Empty;

            TaxonomySession session = TaxonomySession.GetTaxonomySession(ctx);
            TermStore termStore = session.GetDefaultSiteCollectionTermStore();
            TermSetCollection termSets = termStore.GetTermSetsByName("city-NguyenAnhTu", Constants.LCID_ENGLISH);

            ctx.Load(termSets, tsc => tsc.Include(ts => ts.Id));
            ctx.Load(termStore, ts => ts.Id);
            ctx.ExecuteQuery();

            termStoreId = termStore.Id;
            termSetId = termSets.FirstOrDefault().Id;
        }

        private static async Task GetFieldTermValue(ClientContext ctx, string termId)
        {
            //load term by id
            TaxonomySession session = TaxonomySession.GetTaxonomySession(ctx);
            Term taxonomyTerm = session.GetTerm(new Guid(termId));
            ctx.Load(taxonomyTerm, t => t.Labels,
                                   t => t.Name,
                                   t => t.Id);
            await ctx.ExecuteQueryAsync();
        }

        private static async Task ExampleSetTaxonomyFieldValue(ListItem item, ClientContext ctx)
        {
            var field = ctx.Web.Fields.GetByTitle("fieldname");

            ctx.Load(field);
            await ctx.ExecuteQueryAsync();

            var taxField = ctx.CastTo<TaxonomyField>(field);

            taxField.SetFieldValueByValue(item, new TaxonomyFieldValue()
            {
                WssId = -1, // alway let it -1
                Label = "correct label here",
                TermGuid = "term id"
            });
            item.Update();
            await ctx.ExecuteQueryAsync();
        }

        private static async Task CsomTermSetAsync(ClientContext ctx)
        {
            // Get the TaxonomySession
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            // Get the term store by name
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            // Get the term group by Name
            TermGroup termGroup = termStore.Groups.GetByName("CSOM-Test");
            // Get the term set by Name
            TermSet termSet = termGroup.TermSets.GetByName("city-NguyenAnhTu");
            var terms = termSet.GetAllTerms();
            ctx.Load(terms);
            await ctx.ExecuteQueryAsync();
        }

        private static async Task CsomLinqAsync(ClientContext ctx)
        {
            var fieldsQuery = from f in ctx.Web.Fields
                              where f.InternalName == "Test" ||
                                    f.TypeAsString == "TaxonomyFieldTypeMulti" ||
                                    f.TypeAsString == "TaxonomyFieldType"
                              select f;

            var fields = ctx.LoadQuery(fieldsQuery);
            await ctx.ExecuteQueryAsync();
        }

        private static async Task SimpleCamlQueryAsync(ClientContext ctx)
        {
            var list = ctx.Web.Lists.GetByTitle("Documents");

            var allItemsQuery = CamlQuery.CreateAllItemsQuery();
            var allFoldersQuery = CamlQuery.CreateAllFoldersQuery();

            var items = list.GetItems(new CamlQuery()
            {
                ViewXml = @"<View>
                                <Query>
                                    <OrderBy><FieldRef Name='ID' Ascending='False'/></OrderBy>
                                </Query>
                                <RowLimit>20</RowLimit>
                            </View>",
                FolderServerRelativeUrl = "/sites/test-site-duc-11111/Shared%20Documents/2"
                //example for site: https://omniapreprod.sharepoint.com/sites/test-site-duc-11111/
            });

            ctx.Load(items);
            await ctx.ExecuteQueryAsync();
        }

        private static async Task CAMLQueryAsync(ClientContext ctx)
        {
            List targetList = ctx.Web.Lists.GetByTitle("CSOM Test");

            // get list items where field “about” is not “about default”
            // Remmember syntax must be correct
            ListItemCollection oCollection = targetList.GetItems(new CamlQuery()
            {
                ViewXml =
                     @"<View> 
                             <Query> 
                                 <Where><Neq><FieldRef Name='About' /><Value Type='Text'>about default</Value></Neq></Where> 
                             </Query> 
                             <RowLimit>100</RowLimit> 
                     </View>",
            });
            ctx.Load(oCollection);
            await ctx.ExecuteQueryAsync();
            // Print List
            Console.WriteLine("Result of Query: {0}", oCollection.Count());
            foreach (ListItem oItem in oCollection)
            {
                Console.WriteLine("Item : " + oItem["Title"].ToString());
            }
        }

        private static async Task ListViewCSOMOrder(ClientContext ctx)
        {
            List targetList = ctx.Web.Lists.GetByTitle("CSOM Test");
            ViewCollection viewCollection = targetList.Views;
            ctx.Load(viewCollection);
            View listView = viewCollection.Add(new ViewCreationInformation
            {
                Title = "CSOM Order",
                ViewTypeKind = ViewType.Html,
                ViewFields = new String[] { "ID", "Title", "City", "About" },
                Query = "<Where><Eq><FieldRef Name = 'City' /><Value Type = 'Text'>Ho Chi Minh</Value></Eq></Where><OrderBy><FieldRef Name = 'ID' Ascending='FALSE'/></OrderBy>",
            });
            ctx.ExecuteQuery();
            listView.Update();
            await ctx.ExecuteQueryAsync();
        }

        private static async Task UpdateList(ClientContext ctx)
        {
            List targetList = ctx.Web.Lists.GetByTitle("CSOM Test");
            ListItemCollectionPosition itemPosition = null;
            while (true)
            {
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ListItemCollectionPosition = itemPosition;
                camlQuery.ViewXml =
                  @"<View> 
                        <Query> 
                           <Where><Eq><FieldRef Name='About' /><Value Type='Text'>about default</Value></Eq></Where> 
                        </Query> 
                        <RowLimit>2</RowLimit> 
                     </View>";
                ListItemCollection listItems = targetList.GetItems(camlQuery);
                ctx.Load(listItems);
                ctx.ExecuteQuery();
                itemPosition = listItems.ListItemCollectionPosition;
                Console.WriteLine(itemPosition);
                foreach (ListItem listItem in listItems)
                {
                    Console.WriteLine("Item Title: {0}", listItem["Title"]);
                    listItem["About"] = "Update script";
                    listItem.Update();
                    await ctx.ExecuteQueryAsync();
                }
                if (itemPosition == null)
                    break;
            }
        }

        private static async Task AddAuthorToSCOMList(ClientContext ctx)
        {
            //Add Field Author to List
            List targetList = ctx.Web.Lists.GetByTitle("CSOM Test");
            targetList.Fields.AddFieldAsXml("<Field DisplayName='Author' Name='Author' Type='User' />", true, AddFieldOptions.AddFieldInternalNameHint);
            ctx.Load(targetList);
            await ctx.ExecuteQueryAsync();

            //Update Field Author for All List Items
            CamlQuery oQuery = CamlQuery.CreateAllItemsQuery();
            ListItemCollection listItems = targetList.GetItems(oQuery);
            ctx.Load(listItems);
            ctx.ExecuteQuery();
            User admin = ctx.Web.EnsureUser("tu.nguyen.anh@devtusturu.onmicrosoft.com");
            ctx.Load(admin);
            ctx.ExecuteQuery();
            Console.WriteLine(admin.Email);

            //Create a FieldUserValue and set the value to your user
            FieldUserValue userValue = new FieldUserValue();
            userValue.LookupId = admin.Id;

            foreach (ListItem item in listItems)
            {
                Console.WriteLine(item["Title"].ToString());
                item["Author0"] = userValue;
                item.Update();
                ctx.ExecuteQuery();
            }
            await ctx.ExecuteQueryAsync();
        }

        private static async Task CreateCitiesField(ClientContext ctx)
        {
            ctx.Site.RootWeb.Fields.AddFieldAsXml("<Field DisplayName='Cities' Name='Cities' Group='CSOM Test Group' Type='TaxonomyFieldTypeMulti' Mult='TRUE' />", false, AddFieldOptions.AddFieldInternalNameHint);
            await ctx.ExecuteQueryAsync();

            Guid termStoreId = Guid.Empty;
            Guid termSetId = Guid.Empty;
            GetTaxonomyFieldInfo(ctx, out termStoreId, out termSetId);

            Field citiesField = ctx.Site.RootWeb.Fields.GetByInternalNameOrTitle("Cities");
            // Retrieve as Taxonomy Field - Add Term Set to Taxonomy Field
            TaxonomyField taxonomyField = ctx.CastTo<TaxonomyField>(citiesField);
            taxonomyField.SspId = termStoreId;
            taxonomyField.TermSetId = termSetId;
            taxonomyField.TargetTemplate = String.Empty;
            taxonomyField.AnchorId = Guid.Empty;
            taxonomyField.Update();
            await ctx.ExecuteQueryAsync();

            ContentTypeCollection contentTypeCollection = ctx.Web.ContentTypes;
            ctx.Load(contentTypeCollection);
            await ctx.ExecuteQueryAsync();
            ContentType targetContentType = (from contentType in contentTypeCollection where contentType.Name == "CSOM Test Content Type" select contentType).FirstOrDefault();
            ctx.Load(targetContentType);
            await ctx.ExecuteQueryAsync();
            targetContentType.FieldLinks.Add(new FieldLinkCreationInformation
            {
                Field = citiesField
            });
            targetContentType.Update(true);
            await ctx.ExecuteQueryAsync();
        }

        private static async Task AddListWithCitiesField(ClientContext ctx)
        {
            //Get All Terms
            // Get the TaxonomySession
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            // Get the term store by name
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            // Get the term group by Name
            TermGroup termGroup = termStore.Groups.GetByName("CSOM-Test");
            // Get the term set by Name
            TermSet termSet = termGroup.TermSets.GetByName("city-NguyenAnhTu");
            var terms = termSet.GetAllTerms();
            ctx.Load(terms);
            await ctx.ExecuteQueryAsync();
            List targetList = ctx.Web.Lists.GetByTitle("CSOM Test");
            for(int i = 1; i <= 3; i++)
            {
                ListItemCreationInformation iteminfo = new ListItemCreationInformation();
                ListItem newListItem = targetList.AddItem(iteminfo);
                newListItem["Title"] = "Test " + DateTime.Now.ToString();
                foreach(Term term in terms)
                {
                    newListItem["Cities"] = term.Name;
                }
                newListItem.Update();
                ctx.ExecuteQuery();
            }
            await ctx.ExecuteQueryAsync();
        }
    }
}
