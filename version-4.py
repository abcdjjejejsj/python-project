#attendence tracker
import pandas as pd
import uuid
def atten(m):
    global cols
    cols=df.shape[1]
    op={}
    lastC= df.columns[-1]
    lastC= df[lastC].tolist()
    sub={}  #storing name and setting their values to 0
    roll='Roll No.'
    n=df[roll].tolist()
    if(cols<4):
        for i in n:  #assigning value to each key as 0
            key=i
            value=0
            sub[key]=value
    elif(cols==4):
        for i in range(0,68,1):
            if(lastC[i]=="AB"):
                key=n[i]
                value=0
                sub[key]=value
            else:
                key=n[i]
                value=lastC[i]
                sub[key]=value
    else:    #if some data is present
        roll='Roll No.'
        n=df[roll].tolist()
        for i in range(0,68,1):
            row_index=i
            col_index=-1
            cols=df.shape[1]
            if(lastC[i]=="AB"):
                key=n[i]
                prev=lastC[i]
                while(prev=="AB" and cols!=3):
                    prev = df.iloc[row_index, col_index - 1]
                    col_index-=1
                    cols-=1
                if(prev!="AB" and cols==3):
                    prev=0
                value=prev
                sub[key]=value
            else:
                key=n[i]
                value=lastC[i]
                sub[key]=value
    op=sub.copy()    
    temp=sub.copy()
    b=1
    while(b==1):
        date=[]
        c=1
        ab=[]
        temp={}
        k=input("Enter date(dd/mm/yyyy):")
        while(c!=-1): #accepting absent roll numbers
            c=int(input("Enter absent roll number(Enter -1 to stop):"))
            ab.append(c)
        ab.remove(-1)
        for key in sub:
            if key not in ab:
                sub[key]+=1
        for key in sub:
            d=sub[key]
        op=sub.copy()    
        temp=sub.copy()
        for key in temp:
            if key in ab:
                temp[key]="AB"
        for key in temp:
            s=temp[key]
            date.append(s)
        df[k]=date
        print(df)
        if(m==1):
            df.to_excel('PSP.xlsx', index=False) 
        if(m==2):
            df.to_excel('EM-2.xlsx', index=False)
        if(m==3):
            df.to_excel('PE.xlsx', index=False) 
        if(m==4):
            df.to_excel('EP.xlsx', index=False) 
        if(m==5):
            df.to_excel('FDS.xlsx', index=False)
        b=int(input("Press 1 for add new date:"))
def percentage(x,sub):
    if(x==1):
        df = pd.read_excel("PSP.xlsx")
    elif(x==2):
        df = pd.read_excel("EM-2.xlsx")
    elif(x==3):
        df = pd.read_excel("PE.xlsx")
    elif(x==4):
        df = pd.read_excel("EP.xlsx")
    elif(x==5):
        df = pd.read_excel("FDS.xlsx")
    cols=df.shape[1]
    if(cols>3):
        percent = []
        last_column_name = df.columns[-1]
        last_column_data = df[last_column_name].tolist()
        row_index=0
        for i in last_column_data:
            col_index=-1
            nc = df.shape[1]
            if i == "AB":
                prev=i
                while(prev=="AB" and nc!=3):
                    prev = df.iloc[row_index, col_index - 1]  
                    nc-=1
                    col_index-=1
                if(prev!="AB" and nc==3):
                    per=0
                else:
                    per = (prev/(nc - 2))*100
            else:
                per = (i/(nc - 3))*100 
            percent.append(per)
            row_index+=1
        dframe = pd.read_excel("Overall.xlsx")
        dframe[f"{sub},{last_column_name}"] = percent
        dframe.to_excel('Overall.xlsx', index=False)
def overall_attendence():
    dframe = pd.read_excel("Overall.xlsx")
    op = []
    row_index=0
    cols=dframe.shape[1]
    if cols==8:
        last_column_name = dframe.columns[-1]
        last_column_data = dframe[last_column_name].tolist()
        for i in last_column_data:
            count=0
            cols=dframe.shape[1]
            col_index=-1
            prev=i
            while(cols!=3):
                count+=prev
                prev = dframe.iloc[row_index, col_index - 1]
                cols-=1
                col_index-=1
            row_index+=1
            OverPer=count/5
            op.append(OverPer)
        dframe["Overall"]=op
        dframe.to_excel('Overall.xlsx', index=False)  
def copy_o():
    o_df=pd.read_excel('Overall.xlsx')
    s_df=pd.read_excel('sample.xlsx')
    f_name=str(uuid.uuid4())[:4]+'.xlsx' 
    g=o_df
    o_df=s_df
    g.to_excel(f_name,index=False)
    o_df.to_excel("Overall.xlsx",index=False)
    print("Overall attendence stored in ",f_name)      
       
def copy_d():
    o_df=pd.read_excel('low_attendance_students.xlsx')
    s_df=pd.read_excel('sam_de.xlsx')
    f_name=str(uuid.uuid4())[:4]+'.xlsx' 
    g=o_df
    o_df=s_df
    g.to_excel(f_name,index=False)
    o_df.to_excel("low_attendance_students.xlsx",index=False)
    print("detained student list stored in ",f_name) 

def detained():
    df = pd.read_excel('Overall.xlsx')
    # Filter the DataFrame to select rows with overall attendance less than 75%
    filtered_df = df[df['Overall'] < 75]
    # Save the filtered data to a new Excel file
    filtered_df.to_excel('low_attendance_students.xlsx', index=False)

ch=0
while(ch!=7):
    print("Choose the subject:\n1.PSP(Problem Solving Using Python)\n2.EM-2(Engineering Mathematics-2)\n3.PE(Professional English)\n4.EP(Engineering Physics)\n5.FDS(Fundamental Of Data Structure)\n6.Percentage attendence\n7.Exit")
    ch=int(input("Enter choice:"))
    if(ch==1):
        df=pd.read_excel('PSP.xlsx')
        atten(ch)
    elif(ch==2):
        df=pd.read_excel('EM-2.xlsx')
        atten(ch)
    elif(ch==3):
        df=pd.read_excel('PE.xlsx')
        atten(ch)  
    elif(ch==4):
        df=pd.read_excel('EP.xlsx')
        atten(ch)
    elif(ch==5):
        df=pd.read_excel('FDS.xlsx')
        atten(ch)
    elif(ch==6):
        percentage(1,"PSP")
        percentage(2,"EM-2")
        percentage(3,"PE")
        percentage(4,"EP")
        percentage(5,"FDS")
        overall_attendence()
        detained()
        copy_d()
        copy_o()
        print("Excel sheet of Overall attendence and detained students is generated successfully !")
        
    elif(ch==7):
        #exit
        print("Exited Successfully !")
    else:
        print("Wrong Choice")