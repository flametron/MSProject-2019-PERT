# MSProject-2019-PERT
Contains the macro for running PERT Analysis in Microsoft Project 2019

## How to Use

### Seting up the macro
Step 1. Open a New Project inside MS Project 2019  
Step 2. Go to **View > Macros (Dropdown) > Visual Basic**  
![Open Visual Basic](./images/OpenVBFromProject.png)
Step 3. Select **ThisProject(Global.MPT)** From **Microsoft Project Objects** under **ProjectGlobal(Global.MPT)** (We are using global so that we can use this macro on all future projects without having to readd to them)
![Select ThisProject(Global.MPT)](./images/SelectThisProject.png)
Step 4. Paste the code into the codebox that opens (append to the end if some code is already present there)
![Paste Code](./images/PasteTheCode.png)
Step 5. Hit **Save** from either the **toolbox** or Going to **File > Save Global.MPT** or just hit **Ctrl + S**

### Adding the button


## Acknowledgements

Used a bit of code snippet from [This StackOverflow answer by dbmitch](https://stackoverflow.com/a/51144941/8791515)  
Based on [This Microsoft Blog](https://docs.microsoft.com/en-us/archive/blogs/projectified/three-point-estimation-pert-in-project-2010-take-1)
