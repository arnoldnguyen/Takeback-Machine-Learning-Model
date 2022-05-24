from automationBackEnd import *
from tkinter import *


def setupPaths():
    rawDataPath = rawDataEntry.get()
    putInPath = putInEntry.get()
    startRow = int(startEntry.get())
    stopRow = int(stopEntry.get())

    print(rawDataPath)
    print(putInPath)

    try:
        automationRunner(rawDataPath, putInPath, startRow, stopRow)
        Label(root, text="Success!").pack(side=TOP)
    except:
        Label(root, text="Failed!").pack(side=TOP)





if __name__ == "__main__":
    root = Tk()
    root.title("Capsheet Automator")

    label1 = Label(root, text="Enter path of raw data:" )
    label1.pack(side=TOP)

    rawDataEntry = Entry(root, width=15, borderwidth=3)
    rawDataEntry.pack(side=TOP)


    label2 = Label(root, text="Enter path where you want to format rawdata as capsheet:")
    label2.pack(side=TOP)

    putInEntry = Entry(root, width=15, borderwidth=3)
    putInEntry.pack(side=TOP)

    Label(root, text="\n").pack(side=TOP)

    Label(root, text="Enter rawData start row and stop row:").pack(side=TOP)
    Label(root, text="start: ").pack(side=TOP)
    startEntry = Entry(root, width=15, borderwidth=3)
    startEntry.pack(side=TOP)

    stopLabel = Label(root, text="stop:").pack(side=TOP)
    stopEntry = Entry(root, width=15, borderwidth=3)
    stopEntry.pack(side=TOP)

    dimButton = Button(root, text="Enter", padx=10, pady=5, command=setupPaths)
    dimButton.pack(side=TOP)




    root.mainloop()


























#Below are Notes for the openpyxl module -- may be able ot use for allocations prjt

'''
List of useful commands:

op.__version__

type(workbook)

 print(wb.sheetnames)

    sheet = wb['Notes1']

    print(type(sheet))

    print(sheet['A1'].value)

    print(sheet['B1'].value)
    print(sheet['C1'].value)

    #sheet['D1'] = "Hi my name is Sameet Hegde!"
   # wb.save('Practice_openpyxl.xlsx') #can also be used to save in different files

    print(sheet.cell(row=1, column=2).value) #another way ot do this

    print(sheet.max_row)
    print(sheet.max_column)

    print(wb.sheetnames[1])
    print(type(wb.sheetnames[1]))


    if(all(sheetname != "createSheet2" for sheetname in wb.sheetnames)):
        wb.create_sheet(title="createSheet2")
        wb.save("Practice_openpyxl.xlsx")


    folder = os.path.expanduser("~/Desktop/Allocations_Submissions") #os.path.expanduser allows us to use shell path
    os.chdir(folder)
    wb = op.load_workbook('submissions.xlsx')

    print(wb.sheetnames)
    sheet = wb['submissions']

    print(sheet['A1'].value)

    rowlist = list(sheet.rows)


    newWb = op.Workbook()

    ws1 = newWb.create_sheet("Test")

    for cell in rowlist[2]:
        ws1.append()



    newWb.save("Test.xlsx")


'''










