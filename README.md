# MSProject-2019-PERT
Contains the macro for running PERT Analysis in Microsoft Project 2019

## How to Use
Make sure to check out **Setting up the macro** below if you haven't already set this up.  
When you first open a project, **hit the button** to add the necessary Fields (_adds the "Optimistic Duration","Most Likely Duration","Pessimistic Duration","Optimistic Weight","Most Likely Weight","Pessimistic Weight","PERT State"_) beside the "Duration" Field.  
While entering data, **by default the Macro will consider only the weights of the first task as the weights for all tasks**, so you can safely set the other weights (weights from task 2 onwards) to 0.  
Make sure that the **sum of the "Optimistic Weight","Most Likely Weight","Pessimistic Weight" Fields are equal to 6**.  
After you have entered all the data into the Entry Table, **hit the button again** and it will analyze PERT and give the output Duration in the "Duration" field.  
_Raise an issue if you face any problems._  
  
[Sample Project can be found here](./SampleProject.mpp)  
**BE SURE TO ENABLE MACROS WHILE OPENING A PROJECT WITH PERT ANALYSIS REQUIREMENTS IF ASKED, OTHERWISE THE MACRO MIGHT NOW WORK**

### Seting up the macro
Step 1. Open a New Project inside MS Project 2019  
Step 2. Go to **View > Macros (Dropdown) > Visual Basic**  
![Open Visual Basic](./images/OpenVBFromProject.png)  
Step 3. Select **ThisProject(Global.MPT)** From **Microsoft Project Objects** under **ProjectGlobal(Global.MPT)** (We are using global so that we can use this macro on all future projects without having to readd to them)  
![Select ThisProject(Global.MPT)](./images/SelectThisProject.png)  
Step 4. Paste the code from **[the PERTmacro.vba file here](./PERTmacro.vba)** ([Raw can be found here](https://raw.githubusercontent.com/flametron/MSProject-2019-PERT/main/PERTmacro.vba)) into the codebox that opens (append to the end if some code is already present there)  
_if no codebox opens, then right click on **ThisProject(Global.MPT)** and select **View Code**_  
![Paste Code](./images/PasteTheCode.png)  
Step 5. Hit **Save** from either the **toolbox** or Going to **File > Save Global.MPT** or just hit **Ctrl + S** And close the Microsoft Visual Basic for Applications window  

### Adding the button
Step 1. Click on **Customize Quick Access Toolbar** (Top Left Corner of Project Window)
![Customize Quick Access Toolbar](./images/CustomizeQuickAccessBar.png)
Step 2. Go to **More Commands...**  
![Go to More Commands](./images/SelectMoreCommands.png)  
Step 3. Select **Macros** from the **Choose Commands from** dropdown  
![Choose Macros](./images/SelectMacrosFromTheChooseCommandsFromDropdown.png)  
Step 4. Select **PERT** and click on **Add >>**  (Make sure the **Customize Quick Access Toolbar** dropdown is selected to **For all documents**)
![Add PERT](./images/SelectPERTAndClickOnAdd.png)  
Step 5. You Should have **PERT** into the **Quick Access Toolbar _For all documents_**  
![How it should look](./images/YouShouldHavePERTInTheAccessToolBarNow.png)  
Step 6. Hit **OK**  
Step 7. Now you have a button in the **Quick Access Toolbar** (on hovering the tooltip says _PERT_)  
![Final Button](./images/HereIsThePertButton.png)



## Acknowledgements

Used a bit of code snippet from [This StackOverflow answer by dbmitch](https://stackoverflow.com/a/51144941/8791515)  
Based on [This Microsoft Blog](https://docs.microsoft.com/en-us/archive/blogs/projectified/three-point-estimation-pert-in-project-2010-take-1)
