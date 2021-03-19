# FactoryTalk View SE Project Tag Search Application

Generally it is not easy for operator to find instruments/sensors/pumps/motors in which FTV HMI displays contain at a large project,because there are so many PID pages, and operator only navigates to pages through overview button groups or other ways.
This application can create one tool which is based on C++ for merge all XML files to one XML file through extracting tags(parameters), and need to import one VBA program into FTV HMI project. Operator can process one button which link with the VBA program to load XML and search tags to display which pages contain the tags.

Create XML Merge Tool

Use XML_Create_Dialog program which is based on C++ to create XML_Create_Dialog.exe application.

![image](https://user-images.githubusercontent.com/16084196/111804078-b25c3e00-890a-11eb-84b5-2bbed8ec9cb0.png)

The function of this applictaion is to extract selected parameter in FTV HMI pages to one XML file. The format of this XML file like below:

![image](https://user-images.githubusercontent.com/16084196/111803455-04509400-890a-11eb-9fdf-72231b261f41.png)

Of course, first of all, exporting all FTV HMI displays to be XML files is needed.

Import VBA program

Firstly, create one button in HMI page to load operation faceplate, then change the priority of this button to enable VBA Control, and activate Release animation.
![image](https://user-images.githubusercontent.com/16084196/111804736-4f1edb80-890b-11eb-854b-519797f225e2.png)
![image](https://user-images.githubusercontent.com/16084196/111805044-a1f89300-890b-11eb-8898-9164564de2c1.png)


then import VBA program files(FTVSearchTag.frm) into HMI program, and select Reference.
![image](https://user-images.githubusercontent.com/16084196/111805400-f26ff080-890b-11eb-9605-84ef286fe986.png)
![image](https://user-images.githubusercontent.com/16084196/111805301-d9673f80-890b-11eb-8917-cc45f3442609.png)

Add new VBA code into Release animation function.
![image](https://user-images.githubusercontent.com/16084196/111805588-20553500-890c-11eb-9940-8ec7f3eb3d1c.png)

Change the XML file path in VBA code based on different project.
![image](https://user-images.githubusercontent.com/16084196/111805824-5befff00-890c-11eb-93ca-03d90b30724d.png)

If you use custome faceplate not PlantPAx, you can edit below VBA code according to different project.
![image](https://user-images.githubusercontent.com/16084196/111806141-a70a1200-890c-11eb-92af-a7f780250d0d.png)
