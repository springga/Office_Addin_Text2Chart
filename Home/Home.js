/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/
/// <reference path="../App.js" />

// This function is run when the app is ready to start interacting with the host application
// It ensures the DOM is ready before adding click handlers to buttons
Office.initialize = function (reason) {
    $(document).ready(function () {

        // Wire up the click events of the two buttons in the WD_OpenXML_js.html page.
        $('#getOOXMLData').click(function () { convert(); });
        $('#setOOXMLData').click(function () { setOOXML(); });
    });
};

function convert() {
    var report = document.getElementById("status");
    try {
        getOOXML(convertText2Chart);
    } catch (e) {        
        report.innerText = e.message;
    }
}

function getOOXML(callback) {
    // Get a reference to the Div where we will write the status of our operation
    var report = document.getElementById("status");
    var textArea = document.getElementById("dataOOXML");
    Office.context.document.getSelectedDataAsync("ooxml",
        function (result) {
            if (result.status == "succeeded") {
                currentOOXML = result.value;
                report.innerText = "Get OOXML successfully!";
                callback(currentOOXML);
            }
            else {
                throw new Error(result.error.message);
            }
        });
}

function convertText2Chart(currentOOXML) {
    var report = document.getElementById("status");
    var textArea = document.getElementById("dataOOXML");
    if (currentOOXML !== "") {
        parser = new DOMParser();
        xmlDoc = parser.parseFromString(currentOOXML, "text/xml");

        //retrieve paragraphs
        var paragraphs = xmlDoc.getElementsByTagName("w:p");
        var num_nodes = paragraphs.length;
        var results = new Array();
        var tabs = new Array();
        var nodes = new Array();
        var connections = new Array();
        var node_template = getXMLTemplate("Template_node.xml");
        var conn_template = getXMLTemplate("Template_connection.xml");
        var parents = new Array();

        //convert paragraphs to nodes
        for (var i = 0; i < num_nodes; i++) {
            results[i] = paragraphs[i].textContent; //get text content
            if (results[i] != "") {
                tabs[i] = paragraphs[i].getElementsByTagName("w:tab").length;   //count tabs
                //fill nodes id, text
                nodes[i] = node_template.replace("{{id}}", i + 1);
                nodes[i] = nodes[i].replace("{{text}}", results[i]);
                if (i > 0) {
                    if (tabs[i] > tabs[i - 1]) {
                        connections[i] = conn_template.replace("{{srcid}}", i);
                        connections[i] = connections[i].replace("{{destid}}", i + 1);
                        connections[i] = connections[i].replace("{{connid}}", num_nodes + i + 1);
                        parents[tabs[i]] = i;
                    }
                    else if (tabs[i] == tabs[i - 1]) {
                        if (tabs[i] > 0) {
                            connections[i] = conn_template.replace("{{srcid}}", parents[tabs[i]]);
                            connections[i] = connections[i].replace("{{destid}}", i + 1);
                            connections[i] = connections[i].replace("{{connid}}", num_nodes + i + 1);
                        }
                    }
                    else {
                        if (tabs[i] > 0) {
                            connections[i] = conn_template.replace("{{srcid}}", parents[tabs[i]]);
                            connections[i] = connections[i].replace("{{destid}}", i + 1);
                            connections[i] = connections[i].replace("{{connid}}", num_nodes + i + 1);
                        }
                    }
                }
            }            
        }

        var template = getXMLTemplate("Template_document.xml");        
        template = template.replace("{{name}}", "Graph 1");
        template = template.replace("{{nodes}}", nodes.join(""));
        template = template.replace("{{connections}}", connections.join(""));
        textArea.textContent = template;
        report.innerText = "convert successfully!";
    } else {
        throw new Error("Invalid XML!");
    }    
}

function setOOXML() {
    // Get a reference to the Div where we will write the outcome of our operation
    var report = document.getElementById("status");
    var textArea = document.getElementById("dataOOXML");
    currentOOXML = textArea.textContent;
    if (currentOOXML !== "") {
        Office.context.document.setSelectedDataAsync(
            currentOOXML, { coercionType: "ooxml" },
            	function (result) {
                if (result.status == "succeeded") {
                    report.innerText = "The setOOXML function succeeded!";
                }
                else {
                    report.innerText = result.error.message;
                    throw new Error(result.error.message);
                }
            });
    }
    else {
        throw new Error("Invalid XML");
    }
}

function getXMLTemplate(fileName) {
    var myOOXMLRequest = new XMLHttpRequest();
    myOOXMLRequest.open('GET', fileName, false);
    myOOXMLRequest.send();
    if (myOOXMLRequest.status === 200) {
        return myOOXMLRequest.responseText;
    } else {
        throw new Error("Get XML file failed!");
    }
}