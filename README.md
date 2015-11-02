# Office_Addin_Text2Chart

###User Instructions
<br/>
1.Select tab indented text

2.Click 'Get Selected Content and Convert'

3.XML content should be filled in textbox. If not or and an error message shows up, please create an [Issue](https://github.com/springga/Office_Addin_Text2Chart/issues/new)

4.Move the cursor to the place you want to insert the chart

5.Click 'Insert Content from the Markup' to get the hierarchy chart
<br/><br/><br/>

###Technical Guide
<br/>
1.Get/Set document content by getSelectedDataAsync/setSelectedDataAsync

2.Understand OOXML enssentials by [Create better add-ins for Word with Office Open XML](https://msdn.microsoft.com/EN-US/library/office/dn423225.aspx)

3.Parse DOM and get text content of a node by 'textContent' property which is not a standard DOM feature.

4.Prepare OOXML template and place anchors to replace

5.Go through the text structure and convert to nodes and connections.

6.[How to publish in Visual Studio](https://msdn.microsoft.com/en-us/library/dd465337.aspx)
<br/><br/><br/>

###Good Reference:
<br/>
[JavaScript API for Office全景](http://zoom.it/Dhc#full)
