#region License
// Copyright (c) 2013, Joel Serbinski
// All rights reserved.

// Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
// Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
// Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

// THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, 
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS
// BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE 
// GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, 
// STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#endregion

using System;
using System.Data.SqlClient;
using Newtonsoft.Json;


namespace jQueryPlugin_wrappers_csharp
{
    /* --[ dataTables ]--
 * DESC: wrapper class for operating the dataTables plugin
 * 
 * NOTES: classes prefixed with "dt_" mirror our implementation of directly related properties
 * 
 *  
 */
    public class datatables
    {
        public string id_Prefix = "NPU_DT_";
        public string id;
        public dt_options options;

        public datatables()
        {
            this.options = new dt_options();
        }

        public class dt_options
        {
            public string aoColumnDefs = "[ { 'bSearchable': false, 'bVisible': false, 'aTargets': [0] }]"; // 0 would be the id field...
            public string aaSorting = "[]"; // this provides the function of NO DEFAULT SORTING to display data as returned raw from sql/etc.
            public bool bJQueryUI = true; // use jquery ui theme
            public bool bPaginate = true; // paging
            public string sPaginationType = "'full_numbers'"; // for the paging navigation (either full or 2 arrows)
            public string aLengthMenu = "[[25, 50, 100, 200, -1], [25, 50, 100, 200, 'All']]"; // default page amounts
            public int iDisplayLength = 25; // must be set to initial pagination amount!!!
            public string sScrollY = "''"; // Set height to your container size for proper sizing & control
            public string sScrollX = "'100%'"; // Set width to container size...add scrollbar in table
            public string sHeightMatch = "'none'"; // none=do not let calculate row height...for faster display
            public string sAjaxSource= null; // ajax call for data (DO NOT SET IF NOT USING WILL CAUSE ERRORS)
            public bool bSortClasses = false; // keeps #bered column sort names, set on sort
            public bool bStateSave = false; // save a cookie to remember state from page submit
            public string fnInitComplete = //"''" // runs on init complete...
                // The next line adds a row click function (allows selection multiple)
                "function () {$('tr').click(function () {if ($(this).hasClass('row_selected')) $(this).removeClass('row_selected'); else $(this).addClass('row_selected'); }); }"
                ; // 
            public string fnRowCallback = null; // if set, runs on row creation
            public string aaData = ""; // holds all table data in json

            // our custom Special Options (built-in)
            public bool highlighting = false; // default no
            public int columnCount = -1; // if set to greater than 0 we will use this
        }

        public string BuildOutput(ref SqlDataReader datareader, string columnames="")
        {
            string html ="",js="",data="";
            int columns = -1;
            //. build table
            html = "<table cellpadding='0' cellspacing='0' border='0' class='display' id='"+this.id+"'>" +
                "<thead><tr>";            
            if (columnames != "") {
                string[] columname = columnames.Split(new char[] { ',' });
                for (int t = 0; t < columname.Length; t++) {html = html + "<th>" + columname[t] + "</th>";}
            } else {
                for (int t = 0; t < datareader.FieldCount; t++) { html = html + "<th>" + datareader.GetName(t) + "</th>";}
            }
            html = html + "</tr></thead>"+
                "<tbody></tbody></table>";

            //. get data!            
            if (datareader.HasRows)
            {
                while (datareader.Read())
                {
                    if (options.aoColumnDefs == "''") {
                        columns = datareader.FieldCount; } // if var is cleared out, we want all fields
                    else { // we expect key field (default behavior) same as aoColumnDefs
                        columns = datareader.FieldCount-1; 
                    }
                    data = data + "[";
                    for (int t = 0; t < datareader.FieldCount; t++)
                    {
                        data = data + JsonConvert.SerializeObject(datareader[t].ToString()) + ",";
                    }
                    data = data.Remove(data.Length - 1, 1) + "],";
                }
                data = data.Remove(data.Length - 1, 1);
            }
            options.aaData = "["+data+"]"; // push data retrieved to our object
            
            //. build js
            js =
            "<script type='text/javascript'>" +
                "var "+id_Prefix + this.id + " = $('#" + this.id + "').dataTable( {" +
                    "'aoColumnDefs': " + options.aoColumnDefs + "," +
                    "'aaSorting': " + options.aaSorting + "," +
                    "'bJQueryUI': " + options.bJQueryUI.ToString().ToLower() + "," +
                    "'bPaginate': "+options.bPaginate.ToString().ToLower() +","+
                    "'sPaginationType': "+options.sPaginationType+"," +
                    "'aLengthMenu': " + options.aLengthMenu + "," +
                    "'iDisplayLength': "+ options.iDisplayLength + "," +
                    "'sScrollY': " + options.sScrollY + "," +
                    "'sScrollX': " + options.sScrollX + "," +
                    "'sHeightMatch': " + options.sHeightMatch + ","+
                    "'bSortClasses': " + options.bSortClasses.ToString().ToLower() + ","+
                    "'bStateSave': " + options.bStateSave.ToString().ToLower() + "," + 
                    "'fnInitComplete': "+options.fnInitComplete+","
                    ;
            // special actions
            if (options.sAjaxSource != null) { js += "'sAjaxSource': " + options.sAjaxSource + ","; }
            if (options.fnRowCallback != null) { js += "'fnRowCallback': " + options.fnRowCallback + ","; }
            // data
            js += "'aaData': " + options.aaData + "});"; 
            if (options.highlighting)
            {
                if (options.columnCount > 0) { columns = options.columnCount; } // allow user customization here
                js +=
                    "$('td', "+id_Prefix + this.id + ".fnGetNodes()).hover(function () {var iCol = $(this).parent().children().index($(this));" + //"$('td').index(this) % "+columns.ToString()+";"+
                        "var nTrs = "+id_Prefix + this.id + ".fnGetNodes();" +
                        "$('td:nth-child('+(iCol+1)+')', nTrs).addClass('highlighted'); }, function () {$('td.highlighted',  "+id_Prefix + this.id + ".fnGetNodes()).removeClass('highlighted'); " +
                    "});";
            }                      
            js+="</script>";

            return html+js;
        }
    }s
}
