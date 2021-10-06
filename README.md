# Optimisation and automation of a real company's invoice handling and tax reporting system in Excel - Excel Advanced (VBA)


### Table of Contents

* [Introduction](#chapter1)


* [A) Analysing how the Excel was before and what to do to improve it](#chapter2)



* [B) The work done and the final improved Excel sheets](#chapter3)
    * [The main sheet 'FACTURATIONS' - Presentation:](#section_3_1)
    * [The main sheet 'FACTURATIONS' - More detailed explanations (data validation, conditional formatting, automation)](#section_3_2)
        * [1) Data validation - Example for a colunm with date](#section_3_2_1)
        * [2) Data validation - Example for a colunm with numbers](#section_3_2_2)
        * [3) Conditional formatting - Case of a duplicate invoice number](#section_3_2_3)
        * [4) Conditional formatting - Case of invoice paid/ unpaid/ not fully paid/ paid in two times](#section_3_2_4)
        * [5) Conditional formatting - Case of an 'avoir' (credit on the invoice)](#section_3_2_5)
    * [The second sheet 'TVA-tab' - Presentation](#section_3_3)
    * [The second sheet 'TVA-tab' - How it works (macro/vba)](#section_3_4)
    


* [C) The code (VBA/macro)](#chapter4)


* [Conclusion](#chapter5)



## Introduction <a class="anchor" id="chapter1"></a>


The main objective of this project is to practice my Advanced Excel skills (data validation, conditional formatting, Macro/VBA, ...) and also to build a fully automated system to generate tax reports and improve the way of doing things (on Excel) of a real company. 

In this project, I will take the company's Excel sheets, analyse them, look at what I can do to improve them and finally use the Excel functions (e.g. data validation, conditional formatting, ...) and use the VBA code and macro to build a clear, simple and automatic system for the compagny. 


I did this project in February 2019 and it was not meant to be presented/published. Its only purpose was to work and be useful for the company. So I didn't think about the presentation element. I wrote in French in the Excel (because it was for a french compagny), you will see it in the images. I think it is not a problem to understand the work done and I will try to translate when necessary to better understand.

To facilitate the understanding of the context, I will explain here some words (used for example in the names of the sheets):

- Facture = invoice / Facturation = invoicing 
- TVA = VAT in english (Value Added Tax)



## A) Analysing how the Excel was before and what to do to improve it<a class="anchor" id="chapter2"></a>


#### There are essentially two Excel sheets:
   - The main one, 'FACTURATIONS', which contains information about the invoice, like: the date, the amount, the company involved, the taxation, whether it has been paid or not and a note if needed.
   - The second, "TVA", where the tax calculation was done manually (it looked somewhat like a draft). 
   
Below we can see a sample of the complete sheets (I have hidden the names of the companies in the 'AFFAIRES' column for confidentiality reasons):

 
![excel_project1]( https://github.com/TristanT56/Optimisation-and-automation-of-a-real-company-s-invoice-handling-and-tax-reporting-system-in-Excel/blob/main/Images%20for%20Readme%20markdown/output_3_0.jpg)
    
![excel_project1]( https://github.com/TristanT56/Optimisation-and-automation-of-a-real-company-s-invoice-handling-and-tax-reporting-system-in-Excel/blob/main/Images%20for%20Readme%20markdown/output_4_0.jpg)
    

**To comment that we will take the example of the 'Entreprise XXX8' in both tabs:**

In the first tab 'FACTURATIONS' we can see for the 'Entreprise XXX8': 
 - DATE D'ENVOI = date the invoice was sent: '19/02/2018'.
 - N°Facture = invoice number: '10.02.2018'.
 - N°CLIENT = client/compagny number: Useless because it is always empty.
 - AFFAIRES = compagny name (or deal name): 'Enterprise XXX8'.
 - MISSION = a note about the project (only the client knows what it is and he wanted tokeep it): '30%'. 
 - H.T = Pre-Tax amount: '1455,09€'.
 - TVA = taxation (it can be a mix of 5,5% 10% or 20%): '10 & 5,5%'.
 - T.T.C = After-Tax amount (caclulation done in the second sheet): '1595€'.
 - TOTAL TTC MOIS = Monthly total of T.T.C: this cell is almost never used.
 - REGLEMENT part = 'Does the invoice was payed?':
       - Cheq reçu le = date the cheque/money was received: '20/02/2018'
       - Montant TTC = how much money was received: '1595€'.
       - Déposé le = date the cheque/money was put at the bank: None (sometime it is mention butnot always).
       - Crédité le = date the money arrived on the bank account: Never mention.
       - Reste du = if the the money received did not correspond to all the invoice, what there is still to pay?: '0€'
 - Note = often used to say which part of the amount was taxed at 5,5% 10% or 20%: 'TVA 10%= 1 330,79€  ; TVA 5,5%=  124,3€'
 
 
 In the second tab 'TVA', it is like a draft to calculate the taxation. There is not always column name but we can find for the 'Entreprise XXX8':
 - (without column name) the name of the compagny: 'Enterprise XXX8'.
 - (without column name) it is not a date but the invoice number: '10.02.2018'. (sometime it is not there)
 - TTC = after-tax amount calculated just after: '1595,01€'
 - The columns: HT 5,5% / HT 10% / HT 20% / HT 0% = repartition of the taxation on the pre-tax amount.
 - Other manual calculations.

We will reorganise, optimise and automate all this, as it can lead to errors and considerable loss of time. 

------------------------

## B) The work done and the final improved excel sheets<a class="anchor" id="chapter3"></a>

### 1 - The main sheet 'FACTURATIONS' - Presentation: <a class="anchor" id="section_3_1"></a>

(I only display few rows as sample but in reality there is records from 2017 to 2021)

There are cells from A to U so I will post the full view but also zoom in on different parts of this main sheet:

![excel_project1]( https://github.com/TristanT56/Optimisation-and-automation-of-a-real-company-s-invoice-handling-and-tax-reporting-system-in-Excel/blob/main/Images%20for%20Readme%20markdown/output_8_0.jpg)


#### Goal of this sheet:

 - To have all the important information about the invoices my customer sends to other companies in one place. But also to be able to see at a glance whether an invoice has been paid or not, so that it is easier for my client to manage all that.

#### My goal:
 - Make sure that the sheet is as clear, simple and automated as possible and that it is impossible to enter wrong elements in the cells to avoid bugs.

**Let's zoom in of the left part (Cells: A to I):**

    
![excel_project1]( https://github.com/TristanT56/Optimisation-and-automation-of-a-real-company-s-invoice-handling-and-tax-reporting-system-in-Excel/blob/main/Images%20for%20Readme%20markdown/output_11_0.jpg)



**Explanation:**

After discussion with the client, I kept only the important and useful cells:

- Cells from A to D: The first four cells are the same as in the old FACTURATIONS sheet, information about the invoice, but I have added data validation to allow only the date and avoid errors, and conditional formatting to highlight any manual entry errors. We'll see this in detail later.

- Cells from E to I: This is the part where the tax is calculated automatically. The client just has to fill in the pre-tax amount and the 5.5%/10%/20% distribution. Here too I added data validation to allow only numbers and avoid errors, set the data type to currency €, did conditional formatting and also added messages/warnings when clicking on cells.  We will see this in detail later.

- We can also see a big 'Insert Row' button to automatically add a new row in the tab and add all the rules (data validation, conditional formatting,...) on it.


**Let's zoom in of the right part (Cells: J to U):**

    
![excel_project1]( https://github.com/TristanT56/Optimisation-and-automation-of-a-real-company-s-invoice-handling-and-tax-reporting-system-in-Excel/blob/main/Images%20for%20Readme%20markdown/output_14_0.jpg)
    


**Explanation:**

- Cells from J to M: this is the part that allows you to see if the invoice has been paid or not.  We have kept only the first 3 columns of the old INVOICES sheet. I added a data validation to allow only the date and avoid errors, and a conditional formatting to highlight if the invoice has been paid or not (green in 'Montant TTC' = ok, red in 'Reste du' = not ok). We will see this in detail later.

- Cells from N to P: this part is for when the company pays the invoice in two times (which is possible). It has the same rules as the cells from J to M and the amount is added to the total automatically.

- Cells from Q to T: Here we can see the totals: Total pre-tax that has not been paid / Total post-tax that has not been paid / Total pre-tax that has been paid / Total post-tax that has been paid. I have also linked the previous cells with conditional formatting to highlight whether the total has been fully paid or not (green = ok, red = not ok). We will see this in detail later.

- Cell 'note': in the past, this was often used to report the tax distribution but now it is used to add a personal note (like a reminder or something) or often to mention that the invoice is considered an 'avoir' (=a credit on the amount).  When it is an 'avoir', I have added conditional formatting and calculations to take itinto account. We will see this in detail later.

**Let's zoom in of the parts under the tab, the legend:**


![excel_project1]( https://github.com/TristanT56/Optimisation-and-automation-of-a-real-company-s-invoice-handling-and-tax-reporting-system-in-Excel/blob/main/Images%20for%20Readme%20markdown/output_17_0.jpg)


**Explanation:**

On the left we can see the legend of the conditional formatting. The legend is organised in three columns: the colour of the affected cell / where it occurs / meaning.

Translation:
1) red / everywhere / when the cell is not filled in (important: cells in red will not be taken for the automatic calculation of the TVA sheet).
2) red / cells E to H (pre-tax amount and tax distribution) / when the total of the tax distribution is different from the pre-tax amount.
3) blue / cells B, D or U (invoice number, mission or note)/ when there is the letter 'A' in the invoice number, or the word 'avoir' in mission or note.
4) red + white font / cell B (invoice number) / when there is already this invoice number and it is a duplicates.
5) white + green font / cell K ('amount of the paiement TTC') / when the invoice is fully paid.
6) white + red font / cells K and M ('amount of the paiement TTC' and 'reste to pay') / when theinvoice is not fully paid.
7) white + blue strikethrough font / cells J to R / when there is the letter 'A' in the invoice number, or the word 'avoir' in mission or note.
8) green + white font / cells Q and R (Unpaid totals) / when all is fully paid (even if it was in one or two times).
9) red + red font / cells Q and R (Unpaid totals) / when all is not fully paid (even if it was in one or two times).
10) red + blue strikethrough font / cells Q and R (Unpaid totals) / when there is the letter 'A' in the invoice number, or the word 'avoir' in mission or note.
 
On the right is an explanation of how the 'avoir' (credit on the invoice) is considered and its effect on the totals. With my contact information if the client has a problem or does not understand something.
 

**Let's zoom in of the parts under the tab, the sum-up chart:**

    
![excel_project1]( https://github.com/TristanT56/Optimisation-and-automation-of-a-real-company-s-invoice-handling-and-tax-reporting-system-in-Excel/blob/main/Images%20for%20Readme%20markdown/output_20_0.jpg)
    

Finally, for the main sheet, we can see a chart and a tab that summarises the totals (paid/unpaid, pre/post tax) by year.

Visualisation was not an objective for this project, which is why there is only one 'quick' chart like this. We can also see a big 'Rafraichir' button to automatically refresh the tab and the graph.

(Ps: I have hidden the totals for privacy reasons)


### 2 - The main sheet 'FACTURATIONS' - More detailed explanations (data validation, conditional formatting, automation): <a class="anchor" id="section_3_2"></a>

#### 1) Data validation - Example for a colunm with date <a class="anchor" id="section_3_2_1"></a>


![excel_project1]( https://github.com/TristanT56/Optimisation-and-automation-of-a-real-company-s-invoice-handling-and-tax-reporting-system-in-Excel/blob/main/Images%20for%20Readme%20markdown/output_23_0.jpg)


**Comment:** 

The data validation rule here forces the user to put only a date (between 2010 and 2040) with the correct format. If the user puts something other than a date or a date with the wrong format, he/she receives an error message and has to try again. This helps to avoid errors, especially when using macros (vba) (we will see this later).

Ps: All columns with dates have this data validation.

#### 2) Data validation - Example for a colunm with numbers <a class="anchor" id="section_3_2_2"></a>

    
![excel_project1]( https://github.com/TristanT56/Optimisation-and-automation-of-a-real-company-s-invoice-handling-and-tax-reporting-system-in-Excel/blob/main/Images%20for%20Readme%20markdown/output_26_0.jpg)


**Comment:** 

The data validation rule here forces the user to put only a number (between -100 000 and 100 000). If the user puts something else, he/she receives an error message and has to try again. This helps to avoid errors, especially when using macros (vba) (we will see this later).

In addition, when I simply click on the cell, an information message is displayed. This is the case for several columns. For this one, it says "Make sure that the total of TVAs (5.5% / 10% / 20%) is equal to the pre-tax amount (called H.T). 

Ps: All columns with numbers have this data validation.

#### 3) Conditional formatting - Case of a duplicate invoice number <a class="anchor" id="section_3_2_3"></a>

    
![excel_project1]( https://github.com/TristanT56/Optimisation-and-automation-of-a-real-company-s-invoice-handling-and-tax-reporting-system-in-Excel/blob/main/Images%20for%20Readme%20markdown/output_29_0.jpg)


**Comment:** The conditional formatting rule is used here to highlight cells containing an error, which is a duplicate invoice number. 

#### 4) Conditional formatting - Case of invoice paid/ unpaid/ not fully paid/ paid in two times <a class="anchor" id="section_3_2_4"></a>


![excel_project1]( https://github.com/TristanT56/Optimisation-and-automation-of-a-real-company-s-invoice-handling-and-tax-reporting-system-in-Excel/blob/main/Images%20for%20Readme%20markdown/output_32_0.jpg)


**Comment:**

The conditional formatting rule is used here to highlight cells to highlight the fact that the invoice is fully paid / unpaid / paid but not fully and paid in two times. You can see in the images the different cases, but globally when it is not paid = red, when it is paid = green. You can also see in the totals that the cells turn green or red depending on this too. This allows users to better see which bill is not fully paid and what its status is.

#### 5) Conditional formatting - Case of an 'avoir' (credit on the invoice) <a class="anchor" id="section_3_2_5"></a>


![excel_project1]( https://github.com/TristanT56/Optimisation-and-automation-of-a-real-company-s-invoice-handling-and-tax-reporting-system-in-Excel/blob/main/Images%20for%20Readme%20markdown/output_35_0.jpg)


**Comment:**

The conditional formatting rule is used here to help the user deal with the case of an 'avoir' (=credit to the invoice). When the user puts an 'A' in the invoice number, or an 'avoir' in the mission or note cell, the conditional formatting is activated. The invoice number, the mission and the note cells turn blue. Furthermore, the amount in 'Reste du' (= the amount still to be paid) and in the totals, also becomes blue and strikethrough. All this allows users to see directly which line is a 'avoir' and better understand what is happening with this situation.

### 3 - The second sheet 'TVA-tab' - Presentation: <a class="anchor" id="section_3_3"></a>
    
    
![excel_project1]( https://github.com/TristanT56/Optimisation-and-automation-of-a-real-company-s-invoice-handling-and-tax-reporting-system-in-Excel/blob/main/Images%20for%20Readme%20markdown/output_38_0.jpg)


**Goal of this sheet:** 

 - To allow the user to generate a simple and clear sheet of all tax records by month and year of his/her invoices, using a simple button. And after that, the user only has to send a copy to his accountant. At the end, the most important information for the user is the 'Total TVA'. All of this is done to avoid any manual errors that the client might make and to reduce his/her working time. 

**My goal:** 
 - Make everything work automatically with macros/vba and make the tables as clear and simple as possible. I also have to make sure that there are no bugs, but also set up a system that allows the user to debug and reset everything to normal operation easily if the user experiences a bug that I had not seen.

### 4 - The second sheet 'TVA-tab' - How it works (macro/vba): <a class="anchor" id="section_3_4"></a>

To understand how it works, I will create a new row/invoice in the main sheet FACTURATIONS as an example. Let's say I sent the invoice to the company "test" on 09/02/2019 and the company "test" sent me the cheque/money on 10/02/2019 (fully paid, no problem, all green).  See below:

    
![excel_project1]( https://github.com/TristanT56/Optimisation-and-automation-of-a-real-company-s-invoice-handling-and-tax-reporting-system-in-Excel/blob/main/Images%20for%20Readme%20markdown/output_42_0.jpg)


Now let's assume that my accountant needs the February tax reports. So I need to update the TVA-tab sheet to see my new paid invoice in the tab so that I have the correct TVA total with all the tax records to give to my accountant. 

**See the image below to follow my explanation:**

- Step 1: I go to the TVA-tab sheet and click on the orange 'Mise à jour' button (= 'Update').
- Step 2: A window pops up proposing two choices: 'Mise à jour TVA' ou 'Enlever le surlignage'. We will see the second choice later but now we want to click on the first one which means 'Update TVA'. I click on it.
- Step 3: A new window appears and again offers me two choices: 'Tableau Entier' to update all the FACTURATIONS sheet from the first row or 'Depuis la dernière fois !' to update only the last updated row/invoice. Normally I use the second choice because it's faster, but let's say I'm not sure if I made a mistake somewhere, so I click on the first choice (update the whole sheet).
- Step 4: A new window appears asking me if I am sure and telling me that if my FACTURATIONS sheet contains a lot of rows (e.g. mutliple years), it may take a few minutes. I know there are only a few lines in FACTURATIONS, so no problem. I click on the second button 'Yes' .
- Step 5: A final window appears and says 'Updating...', a few seconds later I have the result.

(It sounds long when I'm explaining but in reality it is not, it takes 4 clicks)


**See the image below to follow my explanation:**

    
![excel_project1]( https://github.com/TristanT56/Optimisation-and-automation-of-a-real-company-s-invoice-handling-and-tax-reporting-system-in-Excel/blob/main/Images%20for%20Readme%20markdown/output_44_0.jpg)


![excel_project1]( https://github.com/TristanT56/Optimisation-and-automation-of-a-real-company-s-invoice-handling-and-tax-reporting-system-in-Excel/blob/main/Images%20for%20Readme%20markdown/output_45_0.jpg)
    


**Result:** 

I see my test invoice in the right table (February) and in the right place. The table is sorted by the date I deposited the cheque/money at the bank (column "Déposé le") or if it is empty (because I didn't put the money at the bank yet), the date I received the money (column "Cheq recu le").  The calculations are automatic and I get the totals directly. I can copy this tab and send it to my accountant. It only took me four clicks and a few seconds.

Comment:

The button which we used (which updates all rows/invoices in the FACTURATONS sheet starting from the first row/invoice in the sheet) is there to give the user a way to debug the system if he/she finds a bug that I didn't think of.

**Now let's see the "normal" button that updates only the last rows/invoices that were not already in the TVA-Tab sheet:**

To illustrate this case I have created a new invoice as example. Let's say I sent the invoice to the company "test 2" on 09/02/2019 and the company "test 2" sent me the cheque/money on 01/02/2019 (fully paid, no problem, all green). This time I also mentioned that I put the money in the bank on 05/02/2019 (column 'Déposé le'). See below:

    
![excel_project1]( https://github.com/TristanT56/Optimisation-and-automation-of-a-real-company-s-invoice-handling-and-tax-reporting-system-in-Excel/blob/main/Images%20for%20Readme%20markdown/output_49_0.jpg)
    

Same case scenario: let's assume that my accountant needs the February tax reports. So I need to update the TVA-tab sheet to see my new paid invoice in the tab so that I have the correct TVA total with all the tax records to give to my accountant. 

**See the image below to follow my explanation:**

- Step 1: I go to the TVA-tab sheet and click on the orange 'Mise à jour' button (= 'Update').
- Step 2: A window pops up proposing two choices: 'Mise à jour TVA' ou 'Enlever le surlignage'. I click on the first one which means 'Update TVA'. I click on it.
- Step 3: A new window appears and again offers me two choices: 'Tableau Entier' to update all the FACTURATIONS sheet from the first row or 'Depuis la dernière fois !' to update only the last updated row/invoice. This time I use the second choice.
- Step 4: A final window appears and says 'Updating...', a few seconds later I have the result.

(It sounds long when I'm explaining but in reality it is not, it takes 3 clicks and it is even faster than the previous case)


**See the image below to follow my explanation:**

   
![excel_project1]( https://github.com/TristanT56/Optimisation-and-automation-of-a-real-company-s-invoice-handling-and-tax-reporting-system-in-Excel/blob/main/Images%20for%20Readme%20markdown/output_51_0.jpg)
    
    
![excel_project1]( https://github.com/TristanT56/Optimisation-and-automation-of-a-real-company-s-invoice-handling-and-tax-reporting-system-in-Excel/blob/main/Images%20for%20Readme%20markdown/output_52_0.jpg)
    



**Result:** 

This time I clicked on the button that updates the TVA-tab tables with only the last rows/invoices of the FACTURATIONS sheet that were not yet in the TVA-tab sheet.
With this button, the system highlights the new rows/invoices that we have brought in the tab so that we can see directly where they are and quickly check if everything is good.
On the right side of the table we can also see "Ligne ajoutée avec la dernière mise à jour" which means "Row added with the last update." to emphasize what we said before and therefore better see where the new rows are. 


We can see that the invoice corresponding to the test 2 is in the right table (February) and in the right place. It is placed not with its date we received the money (column "Cheq recu le") but with the date when we deposited the cheque/money at the bank (column "Déposé le"). This date has the priority. 

The calculations are automatic and I get the totals directly. 

**Before sending it to my accountant, I can remove the highlighting by using the last button we haven't used yet:**

    
![excel_project1]( https://github.com/TristanT56/Optimisation-and-automation-of-a-real-company-s-invoice-handling-and-tax-reporting-system-in-Excel/blob/main/Images%20for%20Readme%20markdown/output_54_0.jpg)
    
    
![excel_project1]( https://github.com/TristanT56/Optimisation-and-automation-of-a-real-company-s-invoice-handling-and-tax-reporting-system-in-Excel/blob/main/Images%20for%20Readme%20markdown/output_55_0.jpg)


It's all good now, I can copy it and send it to my accountant.

-----------

## C) The code (VBA/macro)<a class="anchor" id="chapter4"></a>


I'm sorry to say that when I made this project I made a mistake: I didn't think about presenting it when I created it in 2019, it just had to work well to meet the objectives set and so I made the mistake of forgetting to add comments on the code. I now know the importance of doing this and I do it for every new project. For this project, unfortunately I have other projects/priorities to deal with at the moment, so I will just copy and paste the code here without explanation (but I want to do it soon). 

I hope you will forgive me for this and that you will be able to understand some parts of the code:

--

To make this work, I also used hidden cells as references so that the system knows, for example, which rows/invoices were last updated in the TVA-tab sheet.

Here are these cells (you will see them mentioned in the code):

    
![excel_project1]( https://github.com/TristanT56/Optimisation-and-automation-of-a-real-company-s-invoice-handling-and-tax-reporting-system-in-Excel/blob/main/Images%20for%20Readme%20markdown/output_59_0.jpg)
    

**Comment:** as mentioned earlier, these cells are normally hidden and are only there for the code/system to work. Indeed, the functions in the code will look to see if there is a 'Oui' in the V or Z column, which means that a manual change has been made to the row/invoice and therefore that this row/invoice needs to be updated in the TVA-tab. If there is no 'Oui', it means that everything is already in the TVA-tab.

**Next, the full code:**

```
Public Sub DerLigneMAJ()
    Worksheets("FACTURATION-Tab").Select
    trierDate
    
                        Dim MAJPasFait As Range
                        Dim AC As Integer
                        Dim Lastr As Long
                        Dim rangeToSearch As Range
                        Dim i As Integer
                        Dim MAJReste As Integer
                        MAJReste = 0
                        i = 1

                        
    
    Range("V5").Offset(1, 0).Select
                                
                        AC = ActiveCell.Column
                        Lastr = Cells(Rows.Count, AC).End(xlUp).Offset(-1, 0).Row
                        'Debug.Print "derniere ligne= "; Lastr
                        Set rangeToSearch = Range(Cells(ActiveCell.Row, AC), Cells(Lastr, AC))
                        'Debug.Print rangeToSearch.Count
                        
                        Do While i <= rangeToSearch.Count
                            
                            If ActiveCell = "Oui" And Not IsEmpty(ActiveCell.Offset(0, -11)) Then
                                Set MAJPasFait = ActiveCell
                                'Debug.Print ActiveCell.Value
                                'Debug.Print MAJPasFait.Address
                                    Call Lunch(MAJPasFait.EntireRow.Cells(1).Address, MAJReste)
                                    MAJPasFait.Select
                                    
                                
                            ElseIf ActiveCell = "Oui" And IsEmpty(ActiveCell.Offset(0, -11)) Then
                                Set MAJPasFait = ActiveCell
                                findSupMe (MAJReste)
                                
                            End If
                            ActiveCell.Offset(1, 0).Select
                         i = i + 1
                         Loop
                         
                         
Range("Z5").Offset(1, 0).Select
MAJReste = 1
    
                        i = 1
                        AC = ActiveCell.Column
                        Lastr = Cells(Rows.Count, AC).End(xlUp).Offset(-1, 0).Row
                        'Debug.Print "derniere ligne= "; Lastr
                        Set rangeToSearch = Range(Cells(ActiveCell.Row, AC), Cells(Lastr, AC))
                        'Debug.Print rangeToSearch.Count
                        
                        Do While i <= rangeToSearch.Count
                            
                            If ActiveCell = "Oui" And Not IsEmpty(ActiveCell.Offset(0, -11)) Then
                                Set MAJPasFait = ActiveCell
                                'Debug.Print ActiveCell.Value
                                'Debug.Print MAJPasFait.Address
                                    Call Lunch(MAJPasFait.EntireRow.Cells(1).Address, MAJReste)
                                    MAJPasFait.Select
                                    
                                
                            ElseIf (ActiveCell = "Oui" And IsEmpty(ActiveCell.Offset(0, -11))) Then
                                Set MAJPasFait = ActiveCell
                                findSupMe (MAJReste)
                                
                            End If
                            ActiveCell.Offset(1, 0).Select
                         i = i + 1
                         Loop
End Sub


Public Sub Lunch(debut As String, MAJReste As Integer)
    Dim nb As Integer
    If debut = "Oui" Then
        Worksheets("FACTURATION-Tab").Select
        trierDate
        Range("A5").Offset(1, 0).Select
        nb = Range(Selection, Selection.End(xlDown)).Count
    Else:
        nb = 1
    End If
    
    Dim i As Variant
    i = 1
    Do While i <= nb
        If debut = "Oui" Then
        Call Control(debut, MAJReste)
        ActiveCell.EntireRow.Cells(1).Select
        ActiveCell.Offset(1, 0).Select
        Else:
        Call Control(debut, MAJReste)
        End If
        
        
    i = i + 1
    Loop
End Sub


Public Sub findSupMe(MAJReste As Integer)
    Dim NumFact As Variant
    Dim DateMAJ As Variant
    
    
    If MAJReste = 1 Then
        NumFact = ActiveCell.EntireRow.Cells(29)
        DateMAJ = ActiveCell.EntireRow.Cells(27)
    ElseIf MAJReste = 0 Then
        NumFact = ActiveCell.EntireRow.Cells(25)
        DateMAJ = ActiveCell.EntireRow.Cells(23)
    End If
    'Debug.Print "NumFact= " & NumFact
    'Debug.Print "DateMAJ= " & DateMAJ
        If Not IsEmpty(NumFact) Then
    
                        Worksheets("TVA-Tab").Select
                        Dim ici As Range
                        Dim AC As Integer
                        Dim Lastr As Long
                        Dim rangeToSearch As Range
                        
                        Range("A1").Select
                        AC = ActiveCell.Column
                        Lastr = Cells(Rows.Count, AC).End(xlUp).Row
                        Set rangeToSearch = Range(Cells(ActiveCell.Row, AC), Cells(Lastr, AC))
                        Set ici = rangeToSearch.find(NumFact, ActiveCell, , , xlByRows, xlNext)
                        If Not ici Is Nothing Then
                            Do While Not ici Is Nothing
                                
                                ici.Select
                                If ici.EntireRow.Cells(10) = DateMAJ Or ici.EntireRow.Cells(9) = DateMAJ Then
                                    ici.EntireRow.Delete Shift:=xlUp
                                    ModifTotaux
                                    ActiveCell.EntireRow.Cells(1).Select
                                    
                                    
                                End If
                                
                                Set rangeToSearch = Range(Cells(ActiveCell.Row, AC), Cells(Lastr, AC))
                                Set ici = rangeToSearch.find(NumFact, ActiveCell, , , xlByRows, xlNext)
                                
                            Loop
                        End If
            End If
    Worksheets("FACTURATION-Tab").Select
    If MAJReste = 1 Then
        ActiveCell.EntireRow.Cells(26).Select
        ActiveCell.Offset(0, 1).ClearContents
        ActiveCell.Offset(0, 2).ClearContents
        ActiveCell.Offset(0, 3).ClearContents
    ElseIf MAJReste = 0 Then
        ActiveCell.EntireRow.Cells(22).Select
        ActiveCell.Offset(0, 1).ClearContents
        ActiveCell.Offset(0, 2).ClearContents
        ActiveCell.Offset(0, 3).ClearContents
    End If
End Sub

Public Sub Control(debut As String, MAJReste As Integer)
    Dim goPasgo As Integer
    Dim vide As Variant
    vide = 0
    goPasgo = 0
    Dim Avoir As Integer
    Avoir = 0
    
    
    
    If MAJReste = 1 Then
        ActiveCell.EntireRow.Cells(15).Select
    Else: ActiveCell.EntireRow.Cells(11).Select
    End If
    If Not IsEmpty(ActiveCell) Then
        If InStr(ActiveCell.EntireRow.Cells(2), "A") Or InStr(ActiveCell.EntireRow.Cells(2), "a") Then
        'Debug.Print "A ou a dans le numFacture"
        Avoir = 1
        End If
        
        
        If InStr(ActiveCell.EntireRow.Cells(4), "avoir") Or InStr(ActiveCell.EntireRow.Cells(4), "Avoir") Then
        'Debug.Print "Avoir ou avoir dans la MISSION"
        Avoir = 1
        End If
        
        
        If InStr(ActiveCell.EntireRow.Cells(21), "avoir") Or InStr(ActiveCell.EntireRow.Cells(21), "Avoir") Then
        'Debug.Print "Avoir ou avoir dans NOTES"
        Avoir = 1
        End If
    
        If Not Avoir = 1 Then

            'date que recu ou depot
            If (Not (IsEmpty(ActiveCell.Offset(0, -1)) And IsEmpty(ActiveCell.Offset(0, 1)))) And ((ActiveCell.Offset(0, -1).Interior.Color = RGB(255, 0, 0)) And (ActiveCell.Offset(0, 1).Interior.Color = RGB(255, 0, 0))) Then
                    ActiveCell.Offset(0, -1).Interior.Pattern = xlNone
                    ActiveCell.Offset(0, 1).Interior.Pattern = xlNone
                    goPasgo = goPasgo
            ElseIf (IsEmpty(ActiveCell.Offset(0, -1)) And IsEmpty(ActiveCell.Offset(0, 1))) Then
                        ActiveCell.Offset(0, -1).Interior.Color = RGB(255, 0, 0)
                        ActiveCell.Offset(0, 1).Interior.Color = RGB(255, 0, 0)
                        goPasgo = goPasgo + 1
            End If
     
            'dateEnvoie
            If Not (IsEmpty(ActiveCell.EntireRow.Cells(1))) And (ActiveCell.EntireRow.Cells(1).Interior.Color = RGB(255, 0, 0)) Then
                   ActiveCell.EntireRow.Cells(1).Interior.Pattern = xlNone
                   goPasgo = goPasgo
            ElseIf IsEmpty(ActiveCell.EntireRow.Cells(1)) Then
                        ActiveCell.EntireRow.Cells(1).Interior.Color = RGB(255, 0, 0)
                        goPasgo = goPasgo + 1
            End If
     
            'numFacture
            If Not (IsEmpty(ActiveCell.EntireRow.Cells(2))) And (ActiveCell.EntireRow.Cells(2).Interior.Color = RGB(255, 0, 0)) Then
                   ActiveCell.EntireRow.Cells(2).Interior.Pattern = xlNone
                   goPasgo = goPasgo
            ElseIf IsEmpty(ActiveCell.EntireRow.Cells(2)) Then
                        ActiveCell.EntireRow.Cells(2).Interior.Color = RGB(255, 0, 0)
                        goPasgo = goPasgo + 1
            End If
            
            'Affaire
            If Not (IsEmpty(ActiveCell.EntireRow.Cells(3))) And (ActiveCell.EntireRow.Cells(3).Interior.Color = RGB(255, 0, 0)) Then
                   ActiveCell.EntireRow.Cells(3).Interior.Pattern = xlNone
                   goPasgo = goPasgo
            ElseIf IsEmpty(ActiveCell.EntireRow.Cells(3)) Then
                        ActiveCell.EntireRow.Cells(3).Interior.Color = RGB(255, 0, 0)
                        goPasgo = goPasgo + 1
            End If
            
            
            'HT
            If Not (IsEmpty(ActiveCell.EntireRow.Cells(5))) And (ActiveCell.EntireRow.Cells(5).Interior.Color = RGB(255, 0, 0)) Then
                   ActiveCell.EntireRow.Cells(5).Interior.Pattern = xlNone
                   goPasgo = goPasgo
            ElseIf IsEmpty(ActiveCell.EntireRow.Cells(5)) Then
                        ActiveCell.EntireRow.Cells(5).Interior.Color = RGB(255, 0, 0)
                        goPasgo = goPasgo + 1
            End If
            
            'tva
            If Not ((IsEmpty(ActiveCell.EntireRow.Cells(6)) And IsEmpty(ActiveCell.EntireRow.Cells(7)) And IsEmpty(ActiveCell.EntireRow.Cells(8)))) And Not ((ActiveCell.EntireRow.Cells(6).Interior.Color = 16579574) Or (ActiveCell.EntireRow.Cells(7).Interior.Color = 16579574) Or (ActiveCell.EntireRow.Cells(8).Interior.Color = 16579574)) Then
                   ActiveCell.EntireRow.Cells(6).Interior.Color = 16579574
                   ActiveCell.EntireRow.Cells(7).Interior.Color = 16579574
                   ActiveCell.EntireRow.Cells(8).Interior.Color = 16579574
                   goPasgo = goPasgo
            ElseIf (IsEmpty(ActiveCell.EntireRow.Cells(6)) And IsEmpty(ActiveCell.EntireRow.Cells(7)) And IsEmpty(ActiveCell.EntireRow.Cells(8))) Then
                        ActiveCell.EntireRow.Cells(6).Interior.Color = RGB(255, 0, 0)
                        ActiveCell.EntireRow.Cells(7).Interior.Color = RGB(255, 0, 0)
                        ActiveCell.EntireRow.Cells(8).Interior.Color = RGB(255, 0, 0)
                        vide = 1
                        
                        goPasgo = goPasgo + 1
            End If
            
            'si HT diff total tva ET que les cellules tva sont deja "non rens." > reste rouge
            If ((ActiveCell.EntireRow.Cells(6).Value + ActiveCell.EntireRow.Cells(7).Value + ActiveCell.EntireRow.Cells(8).Value) = ActiveCell.EntireRow.Cells(5).Value) And ((ActiveCell.EntireRow.Cells(6).Interior.Color = RGB(255, 0, 0)) And (ActiveCell.EntireRow.Cells(7).Interior.Color = RGB(255, 0, 0)) And (ActiveCell.EntireRow.Cells(8).Interior.Color = RGB(255, 0, 0)) And (ActiveCell.EntireRow.Cells(5).Interior.Color = RGB(255, 0, 0)) And Not (IsEmpty(ActiveCell.EntireRow.Cells(5)))) Then
                        ActiveCell.EntireRow.Cells(6).Interior.Color = 16579574
                        ActiveCell.EntireRow.Cells(7).Interior.Color = 16579574
                        ActiveCell.EntireRow.Cells(8).Interior.Color = 16579574
                        ActiveCell.EntireRow.Cells(5).Interior.Pattern = xlNone
                        goPasgo = goPasgo
            ElseIf (((ActiveCell.EntireRow.Cells(6).Value + ActiveCell.EntireRow.Cells(7).Value + ActiveCell.EntireRow.Cells(8).Value) < (ActiveCell.EntireRow.Cells(5).Value - 0.05)) Or ((ActiveCell.EntireRow.Cells(6).Value + ActiveCell.EntireRow.Cells(7).Value + ActiveCell.EntireRow.Cells(8).Value) > (ActiveCell.EntireRow.Cells(5).Value + 0.05))) And vide <> 1 Then
                        ActiveCell.EntireRow.Cells(6).Interior.Color = RGB(255, 0, 0)
                        ActiveCell.EntireRow.Cells(7).Interior.Color = RGB(255, 0, 0)
                        ActiveCell.EntireRow.Cells(8).Interior.Color = RGB(255, 0, 0)
                        ActiveCell.EntireRow.Cells(5).Interior.Color = RGB(255, 0, 0)
                        goPasgo = goPasgo + 1
            End If
            
            
            'ttc
            If Not (IsEmpty(ActiveCell.EntireRow.Cells(9))) And (ActiveCell.EntireRow.Cells(9).Interior.Color = RGB(255, 0, 0)) Then
                   ActiveCell.EntireRow.Cells(9).Interior.Pattern = xlNone
                   goPasgo = goPasgo
            ElseIf IsEmpty(ActiveCell.EntireRow.Cells(9)) Then
                        ActiveCell.EntireRow.Cells(9).Interior.Color = RGB(255, 0, 0)
                        goPasgo = goPasgo + 1
            End If
            
            
        If goPasgo = 0 Then
            If debut = "Oui" Then
                Call BaseProgTVA(debut, MAJReste)
            Else:
                findSupMe (MAJReste)
                Call BaseProgTVA(debut, MAJReste)
            End If
        End If
        
      Else:
      findSupMe (MAJReste)
      End If
            
    End If
    
    
    
    
End Sub


Public Sub BaseProgTVA(debut As String, MAJReste As Integer)
'debut As String
    
    
    'Dim debut As String
    'debut = "Non"
    'Debug.Print debut
    
    Dim dateEnt As Date
    Dim Annee As Integer
    Dim Mois As String
    Dim NbMois As Integer
    Dim Jour As Integer
    Dim depot As Integer
    depot = 12 + (MAJReste * 4)
    Dim recu As Integer
    recu = 10 + (MAJReste * 4)
    'Debug.print MAJReste
    'Debug.print depot
    'Debug.print recu
           
                If Not IsEmpty(ActiveCell.EntireRow.Cells(depot)) Then
                        ActiveCell.EntireRow.Cells(depot).Select
                        dateEnt = ActiveCell
                        'Debug.print dateEnt
                        Annee = Year(ActiveCell)
                        ''Debug.print "Annee= "; Annee
                        NbMois = month(ActiveCell)
                        'Debug.print "activecellDate= "; ActiveCell
                        'Debug.print "NbMois= "; NbMois
                        Mois = MonthName(NbMois)
                        'Debug.print "Mois= "; Mois
                        Jour = Day(ActiveCell)
                 
                        Call copiecolle(dateEnt, Annee, NbMois, Mois, Jour, debut, MAJReste)
                        Worksheets("FACTURATION-Tab").Select
                        
                    ElseIf Not IsEmpty(ActiveCell.EntireRow.Cells(recu)) Then
                        ActiveCell.EntireRow.Cells(recu).Select
                        dateEnt = ActiveCell
                        'Debug.print dateEnt
                        Annee = Year(ActiveCell)
                        'Debug.print "Annee= "; Annee
                        NbMois = month(ActiveCell)
                        'Debug.print "activecellDate= "; ActiveCell
                        'Debug.print "NbMois= "; NbMois
                        Mois = MonthName(NbMois)
                        'Debug.print "Mois= "; Mois
                        Jour = Day(ActiveCell)
                       
                        Call copiecolle(dateEnt, Annee, NbMois, Mois, Jour, debut, MAJReste)
                        Worksheets("FACTURATION-Tab").Select
                End If
            
    
    
End Sub


Public Sub copiecolle(dateEnt As Date, Annee As Integer, NbMois As Integer, Mois As String, Jour As Integer, debut As String, MAJReste As Integer)
                    
                    'copie des cellules
                    Dim ValFacture As Variant
                    Dim ValAff As String
                    Dim ValMission As String
                    Dim ValHT As Currency
                    Dim ValTVA55 As Currency
                    Dim ValTVA10 As Currency
                    Dim ValTVA20 As Currency
                    Dim ValTTC As Currency
                    Dim ValTTC2 As Currency
                    ValTTC2 = 0
                    Dim ValChrecu As Date
                    Dim ValDepot As Date
                    Dim MontantTTC As Currency
                    Dim ValReste As Currency
                    
                    Dim y As Integer
                    y = 0
                    
                    ValFacture = ActiveCell.EntireRow.Cells(2)
                    ValAff = ActiveCell.EntireRow.Cells(3)
                    ValMission = ActiveCell.EntireRow.Cells(4)
                    ValHT = ActiveCell.EntireRow.Cells(5)
                    ValTVA55 = ActiveCell.EntireRow.Cells(6)
                    ValTVA10 = ActiveCell.EntireRow.Cells(7)
                    ValTVA20 = ActiveCell.EntireRow.Cells(8)
                    ValTTC = ActiveCell.EntireRow.Cells(9)
                    If MAJReste = 1 Then
                        ValTTC2 = ActiveCell.EntireRow.Cells(9) - ActiveCell.EntireRow.Cells(11)
                        ValChrecu = ActiveCell.EntireRow.Cells(14)
                        MontantTTC = ActiveCell.EntireRow.Cells(15)
                        ValDepot = ActiveCell.EntireRow.Cells(16)
                        ValReste = ActiveCell.EntireRow.Cells(18)
                        ActiveCell.EntireRow.Cells(27).Value = dateEnt
                        ActiveCell.EntireRow.Cells(28).Value = MontantTTC
                        ActiveCell.EntireRow.Cells(29).Value = ValFacture

                    ElseIf MAJReste = 0 Then
                        
                        ValChrecu = ActiveCell.EntireRow.Cells(10)
                        MontantTTC = ActiveCell.EntireRow.Cells(11)
                        ValDepot = ActiveCell.EntireRow.Cells(12)
                        ValReste = ActiveCell.EntireRow.Cells(13)
                        ActiveCell.EntireRow.Cells(23).Value = dateEnt
                        ActiveCell.EntireRow.Cells(24).Value = MontantTTC
                        ActiveCell.EntireRow.Cells(25).Value = ValFacture
                    End If
                    
                    Worksheets("TVA-Tab").Select
                    Call RechercheTab(ActiveCell, Annee, NbMois, Mois, y)
                    Call RechercheLigne(Jour)
                   
                    If MAJReste = 1 Then
                        ActiveCell.Value = ValFacture & " (2)"
                    ElseIf MAJReste = 0 Then
                        ActiveCell.Value = ValFacture
                    End If
                    ActiveCell.Offset(0, 1).Value = ValAff
                    If MAJReste = 1 Then
                        ActiveCell.Offset(0, 2).Value = ValMission & " (Restant)"
                    ElseIf MAJReste = 0 Then
                        ActiveCell.Offset(0, 2).Value = ValMission
                    End If
                    If ValChrecu <> "00:00:00" Then
                    ActiveCell.Offset(0, 8).Value = ValChrecu
                    End If
                    If ValDepot <> "00:00:00" Then
                    ActiveCell.Offset(0, 9).Value = ValDepot
                    ActiveCell.Offset(0, 8).Font.Strikethrough = True
                    End If
                    Call calcNouvTVA(ValHT, ValTVA55, ValTVA10, ValTVA20, ValTTC, ValTTC2, MontantTTC, ValReste, debut)
                    
                    
End Sub


Public Sub calcNouvTVA(ValHT As Currency, ValTVA55 As Currency, ValTVA10 As Currency, ValTVA20 As Currency, ValTTC As Currency, ValTTC2 As Currency, MontantTTC As Currency, ValReste As Currency, debut As String)
        
        If (MontantTTC < ValTTC + 1) And (MontantTTC > ValTTC - 1) Then
                    ActiveCell.Offset(0, 3).Value = ValHT
                    ActiveCell.Offset(0, 4).Value = ValTVA55
                    ActiveCell.Offset(0, 5).Value = ValTVA10
                    ActiveCell.Offset(0, 6).Value = ValTVA20
                    ActiveCell.Offset(0, 7).Value = ValTTC
                    ActiveCell.Offset(0, 10).Value = "NON"
                    ActiveCell.Offset(0, 10).HorizontalAlignment = xlRight
                    ActiveCell.Offset(0, 11).Value = "NON"
                    ActiveCell.Offset(0, 11).HorizontalAlignment = xlRight
                    ''Debug.print debut
                    If Not debut = "Oui" Then
                        ActiveCell.Offset(0, 12).Value = "Ligne ajoutée avec la dernière mise à jour."
                        Range(ActiveCell, ActiveCell.Offset(0, 11)).Interior.ThemeColor = xlThemeColorAccent6
                        Range(ActiveCell, ActiveCell.Offset(0, 11)).Interior.TintAndShade = 0.599993896298105
                        
                     End If
                     
                    
        Else:
        'Recalcule le HT, les TVA 5,5% 10% 20% et le TTC si le paiement n'est pas total
            
            Dim PourcentagePaye As Variant
            Dim Pourcentage55 As Variant
            Dim Pourcentage10 As Variant
            Dim Pourcentage20 As Variant
            Dim NouvHT As Currency
            Dim NouvTTC As Currency
            Dim NouvResteHT As Currency
            
            PourcentagePaye = MontantTTC / ValTTC
            'Debug.print PourcentagePaye
            Pourcentage55 = ValTVA55 / ValHT
            Pourcentage10 = ValTVA10 / ValHT
            Pourcentage20 = ValTVA20 / ValHT
            NouvHT = ValHT * PourcentagePaye
            'Debug.print NouvHT
            NouvTTC = (NouvHT * Pourcentage55 * 0.055 + NouvHT * Pourcentage10 * 0.1 + NouvHT * Pourcentage20 * 0.2) + (NouvHT * Pourcentage55 + NouvHT * Pourcentage10 + NouvHT * Pourcentage20)
            'Debug.print NouvTTC
            NouvResteHT = ValReste / (1 + (Pourcentage55 * 0.055 + Pourcentage10 * 0.1 + Pourcentage20 * 0.2))
            

            
            ActiveCell.Offset(0, 3).Value = NouvHT
            ActiveCell.Offset(0, 4).Value = NouvHT * Pourcentage55
            ActiveCell.Offset(0, 4).NumberFormat = "#,##0.00 $"
            ActiveCell.Offset(0, 5).Value = NouvHT * Pourcentage10
            ActiveCell.Offset(0, 5).NumberFormat = "#,##0.00 $"
            ActiveCell.Offset(0, 6).Value = NouvHT * Pourcentage20
            ActiveCell.Offset(0, 6).NumberFormat = "#,##0.00 $"
            ActiveCell.Offset(0, 7).Value = NouvTTC
            ActiveCell.Offset(0, 10).Value = NouvResteHT
            ActiveCell.Offset(0, 11).Value = ValReste
            If Not debut = "Oui" Then
            ActiveCell.Offset(0, 12).Value = "Ligne ajoutée avec la dernière mise à jour."
            Range(ActiveCell, ActiveCell.Offset(0, 11)).Interior.ThemeColor = xlThemeColorAccent6
            Range(ActiveCell, ActiveCell.Offset(0, 11)).Interior.TintAndShade = 0.599993896298105
            End If
            
        
        End If
        
        'Fait la somme pour les totaux
        ModifTotaux


End Sub

Public Sub ModifTotaux()
        
        ActiveCell.EntireRow.Cells(1).End(xlDown).Offset(-1, 0).Select
        If Not Range(Selection, Selection.End(xlUp)).Count <= 3 Then
        
            Dim Firstr As Long
            Firstr = ActiveCell.Offset(-1, 0).End(xlUp).Offset(2, 0).Row
            'Debug.Print Firstr
    
            ActiveCell.End(xlDown).Offset(-1, 3).Select
            ActiveCell.Value = Excel.WorksheetFunction.Sum(Range(Cells(Firstr, ActiveCell.Column), Cells(ActiveCell.Offset(-1, 0).Row, ActiveCell.Column)))
            ActiveCell.Offset(0, 1).Select
            ActiveCell.Value = Excel.WorksheetFunction.Sum(Range(Cells(Firstr, ActiveCell.Column), Cells(ActiveCell.Offset(-1, 0).Row, ActiveCell.Column)))
            ActiveCell.Offset(0, 1).Select
            ActiveCell.Value = Excel.WorksheetFunction.Sum(Range(Cells(Firstr, ActiveCell.Column), Cells(ActiveCell.Offset(-1, 0).Row, ActiveCell.Column)))
            ActiveCell.Offset(0, 1).Select
            ActiveCell.Value = Excel.WorksheetFunction.Sum(Range(Cells(Firstr, ActiveCell.Column), Cells(ActiveCell.Offset(-1, 0).Row, ActiveCell.Column)))
            ActiveCell.Offset(0, 1).Select
            ActiveCell.Value = Excel.WorksheetFunction.Sum(Range(Cells(Firstr, ActiveCell.Column), Cells(ActiveCell.Offset(-1, 0).Row, ActiveCell.Column)))
            Else:
        
            ActiveCell.End(xlDown).Offset(-1, 3).Value = 0
            ActiveCell.End(xlDown).Offset(-1, 4).Value = 0
            ActiveCell.End(xlDown).Offset(-1, 5).Value = 0
            ActiveCell.End(xlDown).Offset(-1, 6).Value = 0
            ActiveCell.End(xlDown).Offset(-1, 7).Value = 0
        End If
        
End Sub


Public Sub RechercheTab(cellule As Variant, Annee As Integer, NbMois As Integer, Mois As String, y As Integer)

                    Dim fAnnee As Range
                    Dim fMois As Range
                    Dim fMois2 As Range
                    Dim f As Range
                    Dim f2 As Integer
                    Dim AC As Integer
                    Dim Lastr As Long
                    Dim rangeToSearch As Range
                    Dim MoisTrouve As Integer
                    MoisTrouve = 0
                    Dim FinBoucle As Integer
                    FinBoucle = 0
                    
                    Range("D1").Select
                    AC = ActiveCell.Column
                    Lastr = Cells(Rows.Count, AC).End(xlUp).Row
                    'Debug.print Lastr
                    Set rangeToSearch = Range(Cells(ActiveCell.Row, AC), Cells(Lastr, AC))
                    Set fMois = rangeToSearch.find(Mois, ActiveCell, , , xlByRows, xlNext)
                    
                If Not fMois Is Nothing Then
                        Do While (Not fMois Is Nothing)
                            fMois.Select
                            If ActiveCell.Offset(0, -2).Value = Annee Then
                                Set f = ActiveCell.Offset(0, -2)
                                MoisTrouve = 1
                            End If
                            ActiveCell.Offset(1, 0).Select
                            Lastr = Cells(Rows.Count, AC).End(xlUp).Row
                            Set rangeToSearch = Range(Cells(ActiveCell.Row, AC), Cells(Lastr, AC))
                            Set fMois = rangeToSearch.find(Mois, ActiveCell, , , xlByRows, xlNext)
                        Loop
                        If MoisTrouve = 1 Then f.Offset(0, -1).Select
                End If
                If (fMois Is Nothing) And (MoisTrouve = 0) Then
                    
                    Range("A1").Select
                    ActiveCell.Offset(1, 0).Select
                    
                    AC = ActiveCell.Column
                    Lastr = Cells(Rows.Count, AC).End(xlUp).Row
                    Set rangeToSearch = Range(Cells(ActiveCell.Row, AC), Cells(Lastr, AC))
                    Set fAnnee = rangeToSearch.find("ANNEE:", ActiveCell, , , xlByRows, xlNext)
                    
                    If fAnnee Is Nothing Then
                        Range("A30000").Select
                        Selection.End(xlUp).Select
                        ActiveCell.Offset(5, 0).Select
                        If y = 1 Then
                        ActiveCell.Offset(6, 0).Select
                        y = 0
                        End If
                       
                        Set f = ActiveCell.Offset(0, 1)
                        f.Select
                        'Debug.Print "f= "; f
                        Call NouvTab(Annee, Mois, f)
                        
                    Else:
                        Do While (Not fAnnee Is Nothing) And (Not FinBoucle = 1)
                            fAnnee.Select
                            Set f = ActiveCell.Offset(0, 1)
                            
                            If f = Annee Then
                                y = 0
                                'f.Select
                                f2 = month("1 " & ActiveCell.Offset(0, 3))
                                'Debug.print "f2= "; f2
                                
                                If f2 = NbMois Then
                                    FinBoucle = 1
                                End If
                                
                                If NbMois < f2 Then
                                    ActiveCell.Offset(-2, 1).Select
                                    'Selection.End(xlUp).Select
                                    'ActiveCell.Offset(2, 0).Select
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Set f = ActiveCell
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    ActiveCell.Offset(2, 0).Select
                                    Call NouvTab(Annee, Mois, f)
                                    FinBoucle = 1
                                End If
                            End If
                            
                            If Annee < f Then
                                If IsEmpty(Selection.End(xlUp)) Then
                                    Selection.End(xlUp).Offset(5, 0).Select
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                 
                                ElseIf (Selection.End(xlUp).End(xlUp).Offset(0, 1)) <> Annee Then
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                Else:
                                    Selection.End(xlUp).Select
                                    ActiveCell.Offset(5, 0).Select
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                    Selection.EntireRow.Insert
                                    Selection.EntireRow.ClearFormats
                                End If
                                Set f = ActiveCell.Offset(0, 1)
                                Call NouvTab(Annee, Mois, f)
                                FinBoucle = 1
                            End If
                            
                            If Annee > f Then
                                y = 1
                            End If
                            
                            'fin des 3if
                            ActiveCell.Offset(1, 0).Select
                            Lastr = Cells(Rows.Count, AC).End(xlUp).Row
                            Set rangeToSearch = Range(Cells(ActiveCell.Row, AC), Cells(Lastr, AC))
                            'Debug.Print Cells(ActiveCell.Row, AC).Address; " et "; Cells(Lastr, AC).Address
                            Set fAnnee = rangeToSearch.find("ANNEE:", ActiveCell, , , xlByRows, xlNext)
                            'Debug.Print fAnnee.Address
                            If fAnnee Is Nothing Then
                                 Range("A30000").Select
                                 Selection.End(xlUp).Select
                                 ActiveCell.Offset(5, 0).Select
                                 If y = 1 Then
                                    ActiveCell.Offset(6, 0).Select
                                    y = 0
                                 End If
                                 Set f = ActiveCell.Offset(0, 1)
                                 f.Select
                                 Call NouvTab(Annee, Mois, f)
                                 f.Offset(0, -1).Select
                            End If
                            If FinBoucle = 1 Then ActiveCell.Offset(-1, 0).Select
                       Loop
                    End If
            End If
End Sub





Public Sub RechercheLigne(Jour)
    Dim i As Integer
    Dim fin As Integer
    fin = 0
    
    
    'Si rien dans le tab
    If Range(Selection, Selection.End(xlDown)).Count <= 4 Then
        ActiveCell.Offset(2, 0).Select
        'ActiveCell.Offset(1, 0).Select
        Selection.EntireRow.Insert
        Selection.EntireRow.ClearFormats
    
    Else:
        Selection.End(xlDown).Offset(-2, 0).Select
    'si val dans tab, trouver sa place
        i = 1
        Dim nbLigne As Integer
        nbLigne = (Range(Selection, Selection.End(xlUp)).Count - 2)
        Do While (i <= nbLigne) And (fin <> 1)
            If Not IsEmpty(ActiveCell.EntireRow.Cells(10)) Then
            ActiveCell.EntireRow.Cells(10).Select
            Else:
            ActiveCell.EntireRow.Cells(9).Select
            End If
            
            If Jour >= Day(ActiveCell) Then
            ActiveCell.EntireRow.Cells(1).Offset(1, 0).Select
            'ActiveCell.Offset(1, -1).Select
            Selection.EntireRow.Insert
            Selection.EntireRow.ClearFormats
            fin = 1
            Else:
               If (Range(Selection, Selection.End(xlUp)).Count - 1) <= 1 Then
                    ActiveCell.EntireRow.Cells(1).Select
                    Selection.EntireRow.Insert
                    Selection.EntireRow.ClearFormats
                    fin = 1
                Else:
                    ActiveCell.Offset(-1, 0).Activate
                End If
            End If
            
        i = i + 1
        Loop
     End If
     
     
     
End Sub









Public Sub NouvTab(a As Integer, m As String, f As Variant)
    
    f.Offset(0, -1).Select
    ActiveCell.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = "ANNEE:"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = a
    ActiveCell.HorizontalAlignment = xlCenter
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = "MOIS:"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = m
    ActiveCell.HorizontalAlignment = xlCenter
    
    ActiveCell.Offset(1, -3).Select
    ActiveCell.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = "N° FACTURE"
    'range(Selection, Selection.End(xlDown)).Select
    'Selection.NumberFormat = "0.00"
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = "AFFAIRE"
    'range(Selection, Selection.End(xlDown)).Select
    'Selection.NumberFormat = "General"
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = "MISSION"
    'range(Selection, Selection.End(xlDown)).Select
    'Selection.NumberFormat = "General"
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = "H.T"
    'range(Selection, Selection.End(xlDown)).Select
    'Selection.NumberFormat = "#,##0.00 $"
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = "TVA 5,5%"
    'range(Selection, Selection.End(xlDown)).Select
    'Selection.NumberFormat = "#,##0.00 $"
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = "TVA 10%"
    'range(Selection, Selection.End(xlDown)).Select
    'Selection.NumberFormat = "#,##0.00 $"
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = "TVA 20%"
    'range(Selection, Selection.End(xlDown)).Select
    'Selection.NumberFormat = "#,##0.00 $"
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = "T.T.C"
    'range(Selection, Selection.End(xlDown)).Select
    'Selection.NumberFormat = "#,##0.00 $"
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = "Cheq recu le"
    'range(Selection, Selection.End(xlDown)).Select
    'Selection.NumberFormat = "m/d/yyyy"
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = "Déposé le"
    'range(Selection, Selection.End(xlDown)).Select
    'Selection.NumberFormat = "m/d/yyyy"
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = "Reste HT"
    'range(Selection, Selection.End(xlDown)).Select
    'Selection.NumberFormat = "#,##0.00 $"
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = "Reste TTC"
    'range(Selection, Selection.End(xlDown)).Select
    'Selection.NumberFormat = "#,##0.00 $"
    
    '--formating
    f.Offset(0, -1).Select
    Selection.Resize(Selection.Rows.Count, Selection.Columns.Count + 3).Select
    'ActiveCell.Offset(-1, -10).Select
    With Selection.Interior
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.799981688894314
    End With
    With Selection.Font
        .Color = -10477568
        .TintAndShade = 0
    End With
    Selection.Font.Size = 11
    Selection.Font.Bold = True
    '------------bordures
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    '--Separation annee et mois
    Selection.Resize(Selection.Rows.Count, Selection.Columns.Count - 2).Select
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    '--format de la ligne du dessous
    ActiveCell.Offset(1, 0).Select
    Selection.Resize(Selection.Rows.Count, Selection.Columns.Count + 11).Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Interior
        .ThemeColor = xlThemeColorAccent1
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
    End With
    Selection.Font.Bold = True
    Selection.Font.Size = 11
    '---totallignes
    
    ActiveCell.Offset(1, 0).Select
    ActiveCell.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = "Total"
    ActiveCell.Offset(0, 3).Value = 0
    ActiveCell.Offset(0, 4).Value = 0
    ActiveCell.Offset(0, 5).Value = 0
    ActiveCell.Offset(0, 6).Value = 0
    ActiveCell.Offset(0, 7).Value = 0
    
    ActiveCell.Offset(1, 0).FormulaR1C1 = "Total TVA  ="
    'calcule tva
    ActiveCell.Offset(1, 4).FormulaR1C1 = "=R[-1]C*5.5%"
    ActiveCell.Offset(1, 4).NumberFormat = "#,##0.00 $"
    ActiveCell.Offset(1, 5).FormulaR1C1 = "=R[-1]C*10%"
    ActiveCell.Offset(1, 5).NumberFormat = "#,##0.00 $"
    ActiveCell.Offset(1, 6).FormulaR1C1 = "=R[-1]C*20%"
    ActiveCell.Offset(1, 6).NumberFormat = "#,##0.00 $"
    'TotalTVA
    ActiveCell.Offset(1, 1).Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[3]:RC[5])"
    Selection.NumberFormat = "#,##0.00 $"
    ActiveCell.HorizontalAlignment = xlCenter
    '---format
    
    ActiveCell.Offset(-1, -1).Select
    Selection.Resize(Selection.Rows.Count, Selection.Columns.Count + 11).Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Interior
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.799981688894314
    End With
    Selection.Font.Bold = True
    '
    ActiveCell.Offset(1, 4).Select
    Selection.Resize(Selection.Rows.Count, Selection.Columns.Count + 2).Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    Selection.Font.Bold = True
    With Selection.Interior
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
    End With
    '
    ActiveCell.Offset(0, -4).Select
    Selection.Resize(Selection.Rows.Count, Selection.Columns.Count + 1).Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    Selection.Font.Bold = True
    With Selection.Interior
        .ThemeColor = xlThemeColorAccent6
    End With
    Selection.Font.Size = 11
    
    ActiveSheet.Columns("A:L").AutoFit
    Selection.Cells(1).Offset(-3, 0).Select
End Sub


Public Sub trierDate()

    ActiveWorkbook.Worksheets("FACTURATION-Tab").ListObjects( _
        "TabRecapFact").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("FACTURATION-Tab").ListObjects( _
        "TabRecapFact").Sort.SortFields.Add Key:=Range( _
        "TabRecapFact[[#Headers],[#Data],[DATE D''ENVOIE]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("FACTURATION-Tab").ListObjects( _
        "TabRecapFact").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Sub


Public Sub nettoyage()
                        Dim nettoye As Range
                        Dim AC As Integer
                        Dim Lastr As Long
                        Dim rangeToSearch As Range
                        Worksheets("TVA-Tab").Select
                        Range("M1").Select
    
                        AC = ActiveCell.Column
                        Lastr = Cells(Rows.Count, AC).End(xlUp).Row
                        Set rangeToSearch = Range(Cells(ActiveCell.Row, AC), Cells(Lastr, AC))
                        Set nettoye = rangeToSearch.find("Ligne ajoutée avec la dernière mise à jour.", ActiveCell, , , xlByRows, xlNext)
                        If Not nettoye Is Nothing Then
                            Do While Not nettoye Is Nothing
                                nettoye.Select
                                nettoye.ClearContents
                                ActiveCell.EntireRow.Interior.Pattern = xlNone
                                ActiveCell.EntireRow.Interior.TintAndShade = 0
                                Set rangeToSearch = Range(Cells(ActiveCell.Row, AC), Cells(Lastr, AC))
                                Set nettoye = rangeToSearch.find("Ligne ajoutée avec la dernière mise à jour.", ActiveCell, , , xlByRows, xlNext)
                            
                            Loop
                        End If
End Sub
```
-----------

## Conclusion <a class="anchor" id="chapter5"></a>
    
This project is not perfect, I am sure that bugs can occur and the overall system can of course be optimised. However, this was not just a classic case study project, it is a real project with a system that has been used for almost 3 years now, every month, by a real professional and it has worked perfectly so far. Indeed, so far he hasn't had any problems or bugs with it, so I'm happy. And most importantly, this person tells me that this Excel saves him between 30 minutes and 1 hour every month. Year after year that's pretty huge! 

In the end, this was my first Excel Advanced project and I'm proud of it.
