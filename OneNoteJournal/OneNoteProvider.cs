using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

using Microsoft.Office.Interop.OneNote;

namespace OneNoteJournal
{
    public class OneNoteProvider
    {
        private readonly Microsoft.Office.Interop.OneNote.Application oneNote;

        //private readonly string oneNoteXMLHierarchy = "";

        /// <summary>
        /// Constructor. Create instance of Microsoft.Office.Interop.OneNote.Application and get XML hierarchy.
        /// </summary>
        public OneNoteProvider()
        {
            this.oneNote = new Microsoft.Office.Interop.OneNote.Application();
        }

        private Microsoft.Office.Interop.OneNote.Application OneNote
        {
            get { return oneNote; }
        }

        /////////////////////////////////////////////////////////////////////////////////////

        private void OpenHierarchy( string strPath, string strRelativeToObjectID, out string strObjectID, CreateFileType createFileType )
        {
            this.oneNote.OpenHierarchy( strPath, strRelativeToObjectID, out strObjectID, createFileType );
        }

        /////////////////////////////////////////////////////////////////////////////////////

        private string GetOneNoteXML( string _strStartID, HierarchyScope _hierarchyScope )
        {
            string strOneNoteXML = "";

            this.oneNote.GetHierarchy( _strStartID, _hierarchyScope, out strOneNoteXML );

            return strOneNoteXML;
        }

        /////////////////////////////////////////////////////////////////////////////////////

        public string GetNotebookIdFromName( string _strName )
        {
            string strOneNoteXML = GetOneNoteXML( null, HierarchyScope.hsNotebooks );

            var xDoc = XDocument.Parse( strOneNoteXML );
            var xNamespace = xDoc.Root.Name.Namespace;
            var ids = from nb in xDoc.Descendants( xNamespace + "Notebook" )
                      where nb.Attribute( "name" ).Value == _strName
                      select nb.Attribute( "ID" ).Value;

            string strId = null;
            if (0 < ids.Count())
            {
                strId = ids.First();
            }

            return strId;
        }

        /////////////////////////////////////////////////////////////////////////////////////

        public string GetSectionGroupIdFromParentIdAndName( string _strParentId, string _strName )
        {
            string strOneNoteXML = GetOneNoteXML( _strParentId, HierarchyScope.hsSections );

            var xDoc = XDocument.Parse( strOneNoteXML );
            var xNamespace = xDoc.Root.Name.Namespace;
            var ids = from nb in xDoc.Descendants( xNamespace + "SectionGroup" )
                      where nb.Attribute( "name" ).Value == _strName
                      select nb.Attribute( "ID" ).Value;

            string strId = ids.First();

            return strId;
        }

        /////////////////////////////////////////////////////////////////////////////////////

        public string GetSectionIdFromParentIdAndName( string _strParentId, string _strName )
        {
            string strOneNoteXML = GetOneNoteXML( _strParentId, HierarchyScope.hsSections );

            var xDoc = XDocument.Parse( strOneNoteXML );
            var xNamespace = xDoc.Root.Name.Namespace;
            var ids = from nb in xDoc.Descendants( xNamespace + "Section" )
                      where nb.Attribute( "name" ).Value == _strName
                      select nb.Attribute( "ID" ).Value;

            string strId = ids.First();

            return strId;
        }

        /////////////////////////////////////////////////////////////////////////////////////

        public string GetPageContentFromId( string _strPageId )
        {
            string strPageXml = "";

            this.oneNote.GetPageContent( _strPageId, out strPageXml );

            return strPageXml;
        }

        /////////////////////////////////////////////////////////////////////////////////////

        public string GetPathFromId( string _strId )
        {
            string strOneNoteXML = GetOneNoteXML( _strId, HierarchyScope.hsSelf );

            var xDoc = XDocument.Parse( strOneNoteXML );
            var xNamespace = xDoc.Root.Name.Namespace;
            var paths = from nb in xDoc.Descendants()
                      where nb.Attribute( "ID" ).Value == _strId
                      select nb.Attribute( "path" ).Value;

            string strPath = paths.First();

            return strPath;
        }

        /////////////////////////////////////////////////////////////////////////////////////

        public string CreateSectionGroupFromParentId( string _strParentId, string _strName )
        {
            string strId = "";

            // Open / Create Section Group
            OpenHierarchy( _strName, _strParentId, out strId, CreateFileType.cftFolder );

            return strId;
        }

        /////////////////////////////////////////////////////////////////////////////////////

        public string CreateSectionFromParentId( string _strParentId, string _strName )
        {
            string strId = "";

            try
            {
                OpenHierarchy( _strName + ".one", _strParentId, out strId, CreateFileType.cftSection );
            }
            catch ( Exception e )
            {
            }

            return strId;
        }

        /////////////////////////////////////////////////////////////////////////////////////

        public string CreatePageFromParentId( string _strParentId, string _strPageName )
        {
            string strId = "";

            // Check of the page already exists!
            strId = GetPageIdFromParentIdAndPageName( _strParentId, _strPageName );

            if ( String.IsNullOrEmpty( strId ) )
            {
                // Create the new Page (with a default name)
                this.oneNote.CreateNewPage( _strParentId, out strId, NewPageStyle.npsBlankPageWithTitle );

                UpdatePageNameFromPageId( strId, _strPageName );
            }

            return strId;
        }

        /////////////////////////////////////////////////////////////////////////////////////

        public string GetPageIdFromParentIdAndPageName( string _strSectionId, string _strPageName )
        {
            string strId = "";

            string strOneNoteXML = GetOneNoteXML( _strSectionId, HierarchyScope.hsPages );

            var xDoc = XDocument.Parse( strOneNoteXML );
            var xNamespace = xDoc.Root.Name.Namespace;
            var strIds = from item in xDoc.Descendants( xNamespace + "Page" )
                        where item.Attribute( "name" ).Value == _strPageName
                        select item.Attribute( "ID" ).Value;

            if ( 0 < strIds.Count() )
            {
                strId = strIds.First();
            }

            return strId;
        }

        /////////////////////////////////////////////////////////////////////////////////////

        public void NavigateToId( string _strPageId )
        {
            this.oneNote.NavigateTo( _strPageId );
        }

        /////////////////////////////////////////////////////////////////////////////////////

        public void UpdatePageNameFromPageId( string _strPageId, string _strPageName )
        {
            string strPageXml = "";
            this.oneNote.GetPageContent( _strPageId, out strPageXml );

            XElement xPageElement = XElement.Parse( strPageXml );

            // Look for the "T" (i.e. Title) element, and update it to the pagename.  This fixes up other aspects of the page
            var xNamespace = xPageElement.GetNamespaceOfPrefix( "one" );
            var xDescendants = from descendants in xPageElement.Descendants( xNamespace + "T" )
                            select descendants;

            XElement xDescendant = xDescendants.First();
            xDescendant.Value = _strPageName;

            this.oneNote.UpdatePageContent( xPageElement.ToString() );
        }

        /////////////////////////////////////////////////////////////////////////////////////

        private string GetRestrictedHierarchyFromId( string _strId, string _strType )
        {
            string strXml;

            string strOneNoteXML = GetOneNoteXML( null, HierarchyScope.hsPages );

            // We first find the entry we are looking for
            var xDoc = XDocument.Parse( strOneNoteXML );
            var xNamespace = xDoc.Root.Name.Namespace;
            var xElements = from item in xDoc.Elements()
                            from page in item.Descendants( xNamespace + _strType )
                        where page.Attribute( "ID" ).Value == _strId
                        select page;

            // When we've found it, we walk upwards eliminating the other children
            XElement xElement = xElements.First();
            XElement xParent = xElement.Parent;
            while ( null != xParent )
            {
                xParent.RemoveNodes();
                xParent.Add( xElement );

                xElement = xParent;
                xParent = xParent.Parent;
            }

            strXml = xElement.ToString();

            return strXml;
        }

        /////////////////////////////////////////////////////////////////////////////////////
    }
}
