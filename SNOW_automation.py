"""
* Automation PSP: script to update risks, security measures, recommendations in Excel/PowerPoint PSP (CCUBE TotalEnergies templates)
* Author: Guillemare Clément
* Date: 08/01/2022
* Version: 1.0
* All right reserved Wavestone CCUBE 
"""

#---------------------------
#          imports
#---------------------------
import tkinter as tk 
import win32com.client
import re

#---------------------------
#          Variables
#---------------------------
urgent_color=6711039
high_color=51455
acceptable_color=5296274
facultative_color=9881640
negligible_color=12632256

#Values from PSP templates for French/English 
PSP_data={
    "worksheets":{
        "EN":['1-Presentation','2-Exec summary','3-Context','4-Architecture','5-Implemented Measures','6-Risk Analysis','7-Action Plan'],
        "FR":['1-Présentation','2-Exec summary','3-Contexte','4-Architecture','5-Mesures appliquées','6-Analyse de Risques',"7-Plan d'Action"]
    },
    "excel_reco_header":{
        "EN":"Recommendation description",
        "FR":"Description des recommandations"
    },
    "excel_sm_header":{
        "EN":"Security measure description",
        "FR":"Description des mesures de sécurité appliquées"
    },
    "ppt_prj_name":{
        "EN":"Project name",
        "FR":"Nom du projet"
    }
}

#---------------------------
#          Variables
#---------------------------

'''
* Class Element : Parent/Generic class that hold description and associated risks
* Can be either a Security Measure or a Recommendation
'''
class Element:
    def __init__(self,description):
        self.myID="ID"
        self.description=description
        self.risks=[]

    #Add risk to the Element's risks tab
    def add_associated_risk(self,risk):
        self.risks.append(risk)

    #Assess if a given risk in the list of Element's risks tab
    def is_associated_risk(self,risk):
        if risk in self.risks:
            return True
        else:
            return False

    #Return list of associated risks as string in the form : R0X,ROY,R0Z,etc
    def get_associated_risk(self):
        str_risks=""
        if len(self.risks)==0:
            return str_risks
        else:
            for risk in self.risks:
                str_risks+=risk.risk_id+", "
            return str_risks[:-2]
    
'''
* Class Recommendation : represent a PSP recommendation (description, priority, associated risks, etc)
* Inherit from Element
'''
class Recommendation(Element):
    def __init__(self,description,priority):
        super().__init__(description)
        self.priority=priority
        
    def __str__(self):
        return f"ID {self.myID}\nDescription {self.description}\nPriority {self.priority}"

'''
* Class SecurityMeasure : represent a PSP security measure (description, associated risks, etc)
* Inherit from Element
'''
class SecurityMeasure(Element):
    def __init__(self,description):
        super().__init__(description)
          
    def __str__(self):
        return f"ID {self.myID}\nDescription {self.description}"

'''
* Class Risk : object that abstract and hold risk information (risk description, impact, potentiality, gravity)
'''
class Risk:
    def __init__(self,risk_id,theme,description,ini_imp,ini_pot,ini_grav,res_imp,res_pot,res_grav):
        self.risk_id = risk_id
        self.theme = theme
        self.description = description
        self.ini_imp=ini_imp
        self.ini_pot=ini_pot
        self.ini_grav=ini_grav
        self.res_imp=res_imp
        self.res_pot=res_pot
        self.res_grav=res_grav
    
    #Used in the Excel Template (Risk analysis worksheet) to provide the list of associated SecurityMeasures / Recommendations
    def get_associated_asString(self,tab):
        res=""
        for elem in tab:
            if elem.is_associated_risk(self):
                res+=f"{elem.myID}: {elem.description}\n"
        return res
    
    #Return the list of associated SecurityMeasures / Recommendations
    def get_associated(self,tab):
        new_tab=[]
        for elem in tab:
            if elem.is_associated_risk(self):
                new_tab.append(elem)
        return new_tab
    
    def __str__(self):
        return f"ID {self.risk_id}\nTheme {self.theme}\nDescription {self.description}\nInitial Impact {self.ini_imp}\nInit Potent. {self.ini_imp}\nInit Grav. {self.ini_grav}\nResid Impact {self.res_imp} \nResid Potent. {self.res_pot}\nResid Grav {self.res_grav}"


'''
* Class PresentationPPT : object that hold ppt slides (risk synthesis, recommendations, risks, exec sum, etc)
'''
class PresentationPPT:
    def __init__(self,slide_risk_synth,slide_recos,slide_sm,slide_intro,slide_context,slide_classif,slide_execSum):
        self.slide_risk_synth=slide_risk_synth
        self.slide_recos=slide_recos
        self.slide_sm=slide_sm 
        self.slide_intro=slide_intro
        self.slide_context=slide_context
        self.slide_classif=slide_classif
        self.slide_execSum=slide_execSum
        self.slide_risks=[]
    
    def add_slide_risk(self,slide):
        self.slide_risks.append(slide)

'''
* Class ProjectPSP : object that project information (project name, project head, etc)
'''
class ProjectPSP:
    def __init__(self,name,head,division,summary,decision,context,hypothesis,availability,confidentiality,integrity,proof,rto,rpo):
        self.name=name
        self.head=head
        self.division=division
        self.summary=summary
        self.decision=decision
        self.context=context
        self.hypothesis=hypothesis
        self.availability=availability
        self.integrity=integrity
        self.confidentiality= confidentiality
        self.proof=proof
        self.rto=rto
        self.rpo=rpo

   
    def __str__(self):
        return f"Project name {self.name}\nHead {self.head}\nDivision {self.division}\nSummary {self.summary}\nDecision {self.decision}\nContext {self.context}\nHypothesis {self.hypothesis}\nAvailability {self.availability}\nIntegrity {self.integrity}\nConfidentiality {self.confidentiality}\nProof {self.proof}\nRTO {self.rto}\nRPO {self.rpo}"


#---------------------------
#          Excel functions to get PSP information from Excel file
#---------------------------

#Open and return workbook object given filename
def openWorkbook(xlapp, xlfile):
    try:        
        xlwb = xlapp.Workbooks(xlfile)            
    except Exception as e:
        try:
            xlwb = xlapp.Workbooks.Open(xlfile)
        except Exception as e:
            print(e)
            xlwb = None                    
    return(xlwb)

#Check if reco/sm already exist and add associated risks if so
def check_new(new_elem,elem_tab,risk):
    is_new=True
    for elem in elem_tab:
        if elem.description == new_elem.description:
            is_new=False
            #update existing reco adding associated risk
            elem.add_associated_risk(risk)
    return is_new

#Find start and stop rows of risk worksheet table
def get_index(ws,lookup):
    
    start_tab=1
    while ws.Cells(start_tab,3).Value !=lookup or start_tab==100:
        start_tab+=1
    end_tab=start_tab+1
    while ws.Cells(end_tab,3).Value !=None or end_tab==100:
        end_tab+=1
    return start_tab+1,end_tab

#Create reco_tab or sm_tab from risk worksheet
def get_elems_from_RXX(ws_RXX,risk,elem_tab,is_reco,language):
    lookup=PSP_data["excel_reco_header"][language]
    if is_reco==False:
        lookup=PSP_data["excel_sm_header"][language]
    
    start,stop=get_index(ws_RXX,lookup)
    for row in range(start,stop):
        #Create new Element object
        
        new_elem=Element(ws_RXX.Cells(row,3).Value)
        
        #Check if element already exist
        is_new=check_new(new_elem,elem_tab,risk)

        if is_new==True:
            
            if is_reco: 
                #Create new recommendation object
                reco_id="REC"+str(len(elem_tab)+1).zfill(2)
                priority=ws_RXX.Cells(row,4).Value
                new_reco=reco_from_elem(new_elem,priority,reco_id,risk)
                
                elem_tab.append(new_reco)
            else:
                #Create SM object
                sm_id="SM"+str(len(elem_tab)+1).zfill(2)
                new_sm=sm_from_elem(new_elem,sm_id,risk)
    
                elem_tab.append(new_sm)
                
           

"""Given an instance of Element, return a new instance of Recommandation"""
def reco_from_elem(elem,priority,reco_id,risk):
    reco=Recommendation(elem.description,priority)
    reco.myID=reco_id
    reco.add_associated_risk(risk)
    return reco
    

"""Given an instance of Element, return a new instance of SecurityMeasure"""
def sm_from_elem(elem,sm_id,risk):
    sm=SecurityMeasure(elem.description)
    sm.myID=sm_id
    sm.add_associated_risk(risk)
    return sm
    
""" Create risk_tab, reco_tab, sm_tab from RXX worksheets """
def get_PSP_risks_inf(wb,reco_tab,sm_tab,risk_tab,language):

    #Browse RXX worksheets
    for ws in wb.Worksheets:
        if re.match("R\d{2}",ws.Name):
            
            #Get risk from RXX worksheet
            risk=Risk(risk_id=str(ws.Name),
                theme=ws.Cells(4,2).Value,
                description=ws.Cells(4,3).Value,
                ini_imp=ws.Cells(4,4).Value,
                ini_pot=ws.Cells(4,5).Value,
                ini_grav=str(ws.Cells(4,6).Value)[4:],
                res_imp=ws.Cells(4,7).Value,
                res_pot=ws.Cells(4,8).Value,
                res_grav=str(ws.Cells(4,9).Value)[4:]
                )
            risk_tab.append(risk)
            
            #Get Recommendations from RXX worksheet
            get_elems_from_RXX(ws,risk,reco_tab,is_reco=True,language=language)
            

            #Get Security Measures from RXX worksheet
            get_elems_from_RXX(ws,risk,sm_tab,is_reco=False,language=language)
            

    return reco_tab,sm_tab,risk_tab 



""" Extract additional PSP information (project name, context, exec sum, hypothesis, etc) and store them in ProjetPSP """
def get_additional_PSP_inf(wb,language):
    ws_presentation=wb.Worksheets(PSP_data["worksheets"][language][0]) 
    name=ws_presentation.Cells(4,4).Value
    head=ws_presentation.Cells(7,4).Value
    division=ws_presentation.Cells(8,4).Value

    ws_exec_sum=wb.Worksheets(PSP_data["worksheets"][language][1]) 
    summary=ws_exec_sum.Cells(4,2).Value
    decision=ws_exec_sum.Cells(7,2).Value

    ws_context=wb.Worksheets(PSP_data["worksheets"][language][2])
    context=ws_context.Cells(2,2).Value
    hypothesis=ws_context.Cells(8,2).Value
    availability=ws_context.Cells(4,4).Value
    integrity=ws_context.Cells(4,5).Value
    confidentiality=ws_context.Cells(4,6).Value
    proof=ws_context.Cells(4,7).Value
    rto=ws_context.Cells(4,9).Value
    rpo=ws_context.Cells(4,10).Value

    project=ProjectPSP(name=name,
    head=head,division=division,summary=summary,
    decision=decision,context=context,hypothesis=hypothesis,
    availability=availability,confidentiality=confidentiality,integrity=integrity,proof=proof,rto=rto,rpo=rpo)

    return project


#---------------------------
#          Excel functions to write risks,security measures, recommendations on Excel file
#---------------------------

""" Check if cell is empty """
def isEmpty(cell):
    if cell.Value is None:
        return True
    else:
        return False


""" Clear excel cells until column reach empty value """
def clean_excel_column(ws,row,column):
    index=0
    while isEmpty(ws.Cells(row+index,column)) == False:
        ws.Cells(row+index,column).Value=None
        index+=1

""" Clear excel table given an origin row and nber of columns to remove """
def clean_excel_table(ws,init_row,columns):
    for column in columns:
        clean_excel_column(ws,init_row,column)

"""Write risks, security measures, recommendations on Excel"""
def update_excel_file(wb,reco_tab,sm_tab,risk_tab,language):
    #Write recommendations
    ws_recos = wb.Worksheets(PSP_data["worksheets"][language][6]) 
    clean_excel_table(ws_recos,init_row=5,columns=[2,3,4,5])
    for index,reco in enumerate(reco_tab):
        ws_recos.Cells(5+index,2).Value=reco.myID
        ws_recos.Cells(5+index,3).Value=reco.get_associated_risk()
        ws_recos.Cells(5+index,4).Value=reco.description
        ws_recos.Cells(5+index,5).Value=reco.priority
    
    #Write Security Measures
    ws_sm = wb.Worksheets(PSP_data["worksheets"][language][4]) 
    clean_excel_table(ws_sm,init_row=5,columns=[2,3,4])
    for index,sm in enumerate(sm_tab):
        ws_sm.Cells(5+index,2).Value=sm.myID
        ws_sm.Cells(5+index,3).Value=sm.get_associated_risk()
        ws_sm.Cells(5+index,4).Value=sm.description
    
    #Write risks
    ws_risks=wb.Worksheets(PSP_data["worksheets"][language][5]) 
    clean_excel_table(ws_risks,init_row=7,columns=[2,3,4,5,6,7,9,10,11])
    for index,risk in enumerate(risk_tab):
        ws_risks.Cells(7+index,2).Value=risk.risk_id
        ws_risks.Cells(7+index,3).Value=risk.theme
        ws_risks.Cells(7+index,4).Value=risk.description
        ws_risks.Cells(7+index,5).Value=risk.get_associated_asString(sm_tab)
        ws_risks.Cells(7+index,6).Value=risk.ini_imp
        ws_risks.Cells(7+index,7).Value=risk.ini_pot
        ws_risks.Cells(7+index,9).Value=risk.get_associated_asString(reco_tab)
        ws_risks.Cells(7+index,10).Value=risk.res_imp
        ws_risks.Cells(7+index,11).Value=risk.res_pot 

#---------------------------
#          PPT functions to write risks,security measures, recommendations and additional project information on .ppt file
#---------------------------

'''Get Cell TextRange value (working only for ppt tables)'''
def get_shape_item(shape,row,column):
    return shape.Table.Cell(row,column).Shape.TextFrame.TextRange

'''Search for a particular slide and return slide if found'''
def search_slide(slides,lookup):
    is_found=False
    
    slide_iterator = iter(slides)
    for slide in slide_iterator:
        try:
            slide.Shapes(lookup)
            is_found=True
            return is_found,slide
        except:
            pass
    return is_found,None

'''Change foreground color of a ppt table cell based on its value'''
def set_color_cell(text,cell):
    if text=="Urgente" or text=="Urgent" or text=="Priority" or text=="Prioritaire":
        cell.Shape.Fill.ForeColor.RGB=urgent_color
    elif text=="Forte" or text=="High" or text=="Arbitration" or text=="Arbitrage":
        cell.Shape.Fill.ForeColor.RGB=high_color
    elif text=="Facultatif" or text=="Optional":
        cell.Shape.Fill.ForeColor.RGB=facultative_color
    elif text=="Acceptable":
        cell.Shape.Fill.ForeColor.RGB=acceptable_color
    elif text=="Mineure" or text=="Negligible":
        cell.Shape.Fill.ForeColor.RGB=negligible_color

'''Change text color of a ppt table cell based on its value'''
def set_font_cell(text,cell):
    #Set as bold
    cell.Shape.TextFrame.TextRange.Font.Bold=True
    #change color
    if text=="Urgente" or text=="Urgent":    
        cell.Shape.TextFrame.TextRange.Font.Color.RGB=urgent_color
    if text=="Forte" or text=="High":   
        cell.Shape.TextFrame.TextRange.Font.Color.RGB=high_color
    if text=="Facultative" or text=="Optional": 
        cell.Shape.TextFrame.TextRange.Font.Color.RGB=facultative_color



'''Add rows to a ppt table based on table size'''
def add_rows(tab,shape):
    #get number of rows for risks table (whithout header)
    nb_rows=shape.Table.Rows.Count-1
    while nb_rows<len(tab):
        #Add a row
        shape.Table.Rows.Add(2)
        #Get nb_rows
        nb_rows=shape.Table.Rows.Count-1

'''Remove and clean ppt table'''
def clean_table(shape):
    #get number of rows to clean 
    nb_rows=shape.Table.Rows.Count
    
    if nb_rows>2:
        for i in range(3,nb_rows+1):
            #Remove rows
            shape.Table.Rows(2).Delete()
        #clean text from remaining row
        for cell in shape.Table.Rows(2).Cells:
            cell.Shape.TextFrame.TextRange.Text=""

'''Clean all tables in RO1 slide (viewed as reference slide for risks)'''
def clean_RO1_slide(pres):
    clean_table(pres.slide_risks[0].Shapes("Recommendations"))
    clean_table(pres.slide_risks[0].Shapes("SecurityMeasures"))
    clean_table(pres.slide_risk_synth.Shapes("Risks"))
    clean_table(pres.slide_recos.Shapes("Recommendations"))
    clean_table(pres.slide_sm.Shapes("SecurityMeasures"))

'''Write risk, associated security measures and recommendations on slide RXX'''
def update_RXX_slide(slide_risk,risk,reco_tab,sm_tab):

    #Add recommendations information
    recos_from_risk=risk.get_associated(reco_tab)
    shape=slide_risk.Shapes("Recommendations")
    add_rows(recos_from_risk,shape)
    for i,reco in enumerate(recos_from_risk):
        get_shape_item(shape,row=2+i,column=1).Text=reco.myID
        get_shape_item(shape,row=2+i,column=2).Text=reco.description
        #Update cell font
        cell_recoID=shape.Table.Cell(i+2,1)
        set_font_cell(reco.priority,cell_recoID)

    #Add SM information
    sm_from_risk=risk.get_associated(sm_tab)
    shape=slide_risk.Shapes("SecurityMeasures")
    add_rows(sm_from_risk,shape)
    for i,sm in enumerate(sm_from_risk):
        get_shape_item(shape,row=2+i,column=1).Text=sm.myID
        get_shape_item(shape,row=2+i,column=2).Text=sm.description
    
    #Add risk information
    shape=slide_risk.Shapes("Risk")
    get_shape_item(shape,row=2,column=1).Text=risk.risk_id
    get_shape_item(shape,row=2,column=2).Text=risk.theme
    get_shape_item(shape,row=2,column=3).Text=risk.description
    shape=slide_risk.Shapes("SecurityMeasures")
    get_shape_item(shape,row=2,column=3).Text=risk.ini_imp
    get_shape_item(shape,row=2,column=4).Text=risk.ini_pot
    get_shape_item(shape,row=2,column=5).Text=risk.ini_grav
    #update initial gravity cell color
    cell_init_grav=shape.Table.Cell(2,5)
    set_color_cell(risk.ini_grav,cell_init_grav)
    shape=slide_risk.Shapes("Recommendations")
    get_shape_item(shape,row=2,column=3).Text=risk.res_imp
    get_shape_item(shape,row=2,column=4).Text=risk.res_pot
    get_shape_item(shape,row=2,column=5).Text=risk.res_grav
    #update resid gravity cell color
    cell_res_grav=shape.Table.Cell(2,5)
    set_color_cell(risk.res_grav,cell_res_grav)

"""Write risks on risk synthesis slide"""
def update_risks_synth_slide(pres, risk_tab):
    shape=pres.slide_risk_synth.Shapes("Risks")
    add_rows(risk_tab,shape)
    for index,risk in enumerate(risk_tab):
        get_shape_item(shape,row=index+2,column=1).Text=risk.risk_id
        get_shape_item(shape,row=index+2,column=2).Text=risk.theme
        get_shape_item(shape,row=index+2,column=3).Text=risk.description
        get_shape_item(shape,row=index+2,column=4).Text=risk.ini_grav
        get_shape_item(shape,row=index+2,column=5).Text=risk.res_grav
        #update initial gravity cell color
        cell_init_grav=shape.Table.Cell(index+2,4)
        set_color_cell(risk.ini_grav,cell_init_grav)
        #update resid gravity cell color
        cell_res_grav=shape.Table.Cell(index+2,5)
        set_color_cell(risk.res_grav,cell_res_grav)

"""Write recommendations on recommendations slide"""
def update_recos_synth_slide(pres,reco_tab):
    shape=pres.slide_recos.Shapes("Recommendations")
    add_rows(reco_tab,shape)
    for index,reco in enumerate(reco_tab):
        get_shape_item(shape,row=index+2,column=1).Text=reco.myID
        get_shape_item(shape,row=index+2,column=2).Text=reco.get_associated_risk()
        get_shape_item(shape,row=index+2,column=3).Text=reco.description
        get_shape_item(shape,row=index+2,column=7).Text=reco.priority
        #update cell color
        cell_priority=shape.Table.Cell(index+2,7)
        set_color_cell(reco.priority,cell_priority)
        
        #Update cell font
        cell_recoID=shape.Table.Cell(index+2,1)
        set_font_cell(reco.priority,cell_recoID)

"""Write security measures on securityMeasures slide"""
def update_sm_synth_slide(pres,sm_tab):
    shape=pres.slide_sm.Shapes("SecurityMeasures")
    add_rows(sm_tab,shape)
    for index,sm in enumerate(sm_tab):
        get_shape_item(shape,row=index+2,column=1).Text=sm.myID
        get_shape_item(shape,row=index+2,column=2).Text=sm.description

"""Get TextRange from Text shape (to manipulate ppt text objects)"""
def get_textFrame(shape):
    return shape.TextFrame.TextRange

"""Write additional information (project name, context, exec sum, etc) based on ProjectPSP"""
def update_addit_inf_slides(pres,project_inf,language):
    
    #Update project name
    shape=pres.slide_intro.Shapes("NOMPROJET")
    get_textFrame(shape).Text=re.sub("(?:\[{}\])".format(PSP_data["ppt_prj_name"][language]), project_inf.name,get_textFrame(shape).Text)
    shape=pres.slide_context.Shapes("PRJ NAME")
    get_textFrame(shape).Text=project_inf.name
    
    #Update project head and division
    shape=pres.slide_intro.Shapes("CPI")
    get_textFrame(shape).Text=re.sub('(?:\<CPI\>)', project_inf.head,get_textFrame(shape).Text)
    get_textFrame(shape).Text=re.sub('(?:\<Division\>)', project_inf.division,get_textFrame(shape).Text)

    #Update Context
    shape=pres.slide_context.Shapes("CONTEXT")
    get_shape_item(shape,row=2,column=1).Text=project_inf.context

    #Update Hypothesis
    shape=pres.slide_classif.Shapes("Assumptions")
    get_textFrame(shape).Text=project_inf.hypothesis

    #Update DICP
    shape=pres.slide_classif.Shapes("DICP")
    get_shape_item(shape,row=2,column=1).Text=project_inf.availability
    get_shape_item(shape,row=2,column=2).Text=project_inf.integrity
    get_shape_item(shape,row=2,column=3).Text=project_inf.confidentiality
    get_shape_item(shape,row=2,column=4).Text=project_inf.proof

    #Update RTO/RPO
    shape=pres.slide_classif.Shapes("RTO RPO")
    get_shape_item(shape,row=2,column=1).Text=project_inf.rto
    get_shape_item(shape,row=2,column=2).Text=project_inf.rpo

    #Update Exec Sum
    shape=pres.slide_execSum.Shapes("Summary")
    get_textFrame(shape).Text=project_inf.summary
    shape=pres.slide_execSum.Shapes("Decision")
    get_textFrame(shape).Text=project_inf.decision

"""Global function to update ppt with risks, security measures, recommendations and additional project information"""
def update_ppt_file(reco_tab,sm_tab,risk_tab,project_inf,ppt_filename,language):
    #get ppt instance
    PPTApp = win32com.client.GetActiveObject("PowerPoint.Application")
    #get ref to the presentation powerpoint object
    PPTPres=PPTApp.Presentations(ppt_filename)
    
    #[1] Find slides in presentation
    is_found,slide_risk_synth=search_slide(slides=PPTPres.Slides,lookup="Title Risks")
    is_found,slide_R01=search_slide(slides=PPTPres.Slides,lookup="Title Risk")
    is_found,slide_recos=search_slide(slides=PPTPres.Slides,lookup="Title Recommendations")
    is_found,slide_sm=search_slide(slides=PPTPres.Slides,lookup="Title SecurityMeasures")

    #[2] Find additional slides (introduction, context, classification,exec sum)
    is_found,slide_intro=search_slide(slides=PPTPres.Slides,lookup="NOMPROJET")
    is_found,slide_context=search_slide(slides=PPTPres.Slides,lookup="Title Context")
    is_found,slide_classif=search_slide(slides=PPTPres.Slides,lookup="Title Classification")
    is_found,slide_execSum=search_slide(slides=PPTPres.Slides,lookup="Title ExecSum")

    pres=PresentationPPT(slide_risk_synth,slide_recos,slide_sm,slide_intro,slide_context,slide_classif,slide_execSum)
    pres.add_slide_risk(slide_R01)
    
    #[3] Clean tables in RO1 slide
    clean_RO1_slide(pres)
    
    #[4] Duplicate Slide R01 risk template and append new risk slides in the slide risk tab
    for i in range(len(risk_tab)-1):
        new_slide_risk=pres.slide_risks[i].Duplicate()
        pres.add_slide_risk(new_slide_risk)
    
    #[5] Update slides RXX
    for index,risk in enumerate(risk_tab):
        slide_risk=pres.slide_risks[index]
        update_RXX_slide(slide_risk,risk,reco_tab,sm_tab)
    
    #[6] Update risks synthesis slide
    update_risks_synth_slide(pres, risk_tab)
        
    #[7] Update recommendations synthesis slide
    update_recos_synth_slide(pres,reco_tab)
    
    #[8] Update SM synthesis slide
    update_sm_synth_slide(pres,sm_tab)
    
    #[9] Update additional slides
    update_addit_inf_slides(pres,project_inf,language)



#---------------------------
#          Tkinter GUI functions or callback functions
#---------------------------

"""
Controller: Callback function called when user clicks on RUN.
Perform actions (get from excel, write on excel, write on ppt) based on user inputs 
"""
def controller():
    risk_tab=[]
    reco_tab=[]
    sm_tab=[]
    project_inf=None

    language=language_button.config('text')[-1]

    

    #Excel Manipulation
    try:
        #Load excel application
        excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
        wb = openWorkbook(excel, excel_filename.get())
        excel.Visible = True
        
        get_PSP_risks_inf(wb,reco_tab,sm_tab,risk_tab,language)
        project_inf=get_additional_PSP_inf(wb,language)
        

        if action_updateExcel.get():
            update_excel_file(wb,reco_tab,sm_tab,risk_tab,language)
 
    except Exception as e:
        print(e)

    finally:
        # RELEASES RESOURCES
        ws = None
        wb = None
        excel = None
    
    #PowerPoint Manipulation
    if action_updatePPT.get():
        update_ppt_file(reco_tab,sm_tab,risk_tab,project_inf,ppt_filename=ppt_filename.get(),language=language)

       

"""Handle text switch on language button """
def toggle_language_button():
    if language_button.config('text')[-1] == "FR":
        language_button.config(text="EN")
    else:
        language_button.config(text="FR")

#---------------------------
#          MAIN
#---------------------------
if __name__=="__main__":
    
    
    '''
        TKINTER GUI instance initiated
    '''
    root = tk.Tk()
    #title
    root.title("Automation PSP")
    

    root.resizable(0, 0)
    
    # configure the grid
    root.columnconfigure(0, weight=3)
    root.columnconfigure(1, weight=1)

    #Language switch
    language_button=tk.Button(root, width=1, text="EN", command=toggle_language_button)
    language_button.grid(column=0, row=0, sticky=tk.EW, padx=5, pady=5)
    #Filename titles
    tk.Label(root, text='excel filename (with extension):').grid(column=0, row=1, sticky=tk.EW, padx=5, pady=5)
    tk.Label(root, text='ppt filename (with extension):').grid(column=0, row=2, sticky=tk.EW, padx=5, pady=5)

    #Filename input box
    excel_filename = tk.Entry(root)
    ppt_filename = tk.Entry(root)
    excel_filename.grid(row=1, column=1,sticky=tk.W, padx=15, pady=15)
    ppt_filename.grid(row=2, column=1, sticky=tk.W,padx=15, pady=15)

    #Action buttons
    action_updateExcel = tk.BooleanVar()
    tk.Checkbutton(root, text='Update excel document', variable=action_updateExcel).grid(row=3, column=0, sticky=tk.W, padx=15, pady=5)
    action_updatePPT = tk.BooleanVar()
    tk.Checkbutton(root, text='Update ppt document', variable=action_updatePPT).grid(row=4,column=0, sticky=tk.W, padx=15, pady=5)

    #RUN button
    tk.Button(root, text ="Run", command = controller).grid(row=5, column=1, padx=15, pady=15,sticky=tk.EW)
    
    root.mainloop()