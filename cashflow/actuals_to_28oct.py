import openpyxl, copy
wb=openpyxl.load_workbook('out/base.xlsx')
cf=wb['Cash Flow']
GAP=456746.27   # other accounts, from 5 Oct when cheque + savings was $358,696.23
def put(col, rows):
    for r,v in rows.items():
        cf[f'{col}{r}']=v
        cf[f'{col}{r}']._style=copy.copy(cf[f'C{r}']._style)
# ---- week 2: 8 - 14 Oct (Xero cash P&L x 1.1 for GST, payroll none) ----
put('E',{13:'=ROUND(369617.53*1.1,2)',14:0,19:0,20:0,21:0,22:0,23:0,
 27:'=ROUND(28736.64*1.1,2)+39033.02',28:'=ROUND((8125.22+520.4)*1.1,2)',29:'=ROUND((11867.45+2204.89+222.68)*1.1,2)',
 30:'=ROUND(19053.3*1.1,2)',31:'=ROUND((1115.04+560.9+46.69)*1.1,2)',
 35:0,36:0,37:0,38:0,39:'=ROUND(629.85*1.1,2)',40:'=804.11+136',41:0,42:0,43:245,44:0,45:0,46:0,47:0,48:0,52:0,
 59:f'=1036314.34-{GAP}'})
cf['E8']=1
# ---- week 3: 15 - 21 Oct ----
put('G',{13:'=ROUND(103182.08*1.1,2)',14:0,19:'=ROUND(111704.13*Assumptions!$B$10,2)',20:14214.47,21:0,22:0,23:0,
 27:'=ROUND(80501.16*1.1,2)',28:'=ROUND((9090+1628.4)*1.1,2)',29:'=ROUND((4167.5+42.5+9.73)*1.1,2)',30:0,
 31:'=ROUND((2525.5+2214.76+1132+834.55)*1.1,2)',
 35:0,36:'=ROUND(9178.55*1.1,2)',37:'=ROUND(4250*1.1,2)',38:0,39:'=ROUND(3154.18*1.1,2)',40:'=ROUND(92*1.1,2)',41:'=ROUND(906*1.1,2)',42:0,
 43:'=ROUND((44.04+31.82+27.93+12.95+6.1)*1.1,2)',44:0,45:'=ROUND(139.89*1.1,2)+36.86',46:0,47:0,48:0,52:0,
 59:f'=912228.18-{GAP}'})
cf['G8']=1
# ---- week 4: 22 - 28 Oct ----
put('I',{13:'=ROUND(17509*1.1,2)',14:0,19:0,20:0,21:0,22:0,23:0,
 27:'=ROUND(11382*1.1,2)',28:'=ROUND(3055*1.1,2)',29:'=ROUND((2630+43.75)*1.1,2)',30:0,31:'=ROUND((658+361.82)*1.1,2)',
 35:0,36:0,37:'=ROUND(4385*1.1,2)',38:342,39:0,40:0,41:0,42:0,43:0,44:0,45:0,46:0,47:'=ROUND(360*1.1,2)',48:0,52:0,
 59:f'=896676.95-{GAP}'})
cf['I8']=1
cf['K8']=0
notes={
 'O8':'Weeks 1 to 4 are done (flag 1). Weeks 2 to 4 are built off the Xero cash P&L, grossed up for GST, with the bank balance worked back from Xero total cash. Week 5 runs to 4 Nov and is in the 29 Oct forecast file.',
 'O13':'Week 2: Aware $225,887 landed Tue 13 Oct, plus Deloitte, CEO Institute, ASX. Week 3: CBA Sep contract invoices (15 Oct), Aware Appspace $23,317, ICC, Bankwest. Week 4: Aware maintenance $12,100 and the Perth USB-C job (27 Oct), RSNSW. All from Xero cash P&L x 1.1.',
 'O19':'Payroll tab - gross x net %. Pay runs Tue 6 Oct (week 1, $79.7k from the bank) and Tue 20 Oct (week 3, $111,704 gross per Xero, net at 78.8% - swap in the ABA figure). Next one Tue 3 Nov is in the new file.',
 'O20':'Gross x 12%. Actuals are the super in the Xero pay run journals ($13,084 on 7 Oct, $14,214 on 21 Oct).',
 'O27':'139k materials paid 29 Sep, labour $39,033 released and paid 8 Oct. Week 3 actual $88.6k is the equipment purchases in Xero cash P&L 15-21 Oct - a lot bigger than the estimate, worth a look at what it was.',
 'O42':'Brookvale $66,239 was in the week 2 forecast but has not been paid - still open in Xero at 29 Oct, due 30 Oct. It is in the 2 Nov creditors run in the new file.',
 'O46':'~$25k a month to the PL and DL cards. Weeks 2 to 4 actuals are nil here because the card spend is already inside the category lines (Xero cash P&L includes what the cards paid). Only the $10k transfer on 7 Oct is shown, in week 1.',
 'O48':'Creditors Run tab. Week 1 actual is the 1 Oct run ($77,303). No run in weeks 2 to 4 - bills were paid one by one and sit in their category lines.',
 'O52':'="Q1 BAS est. "&TEXT(Assumptions!$B$22+Assumptions!$B$23+Assumptions!$B$24,"$#,##0")&" was in week 4 (28 Oct) but did not go out - Xero bank is flat 28 to 29 Oct. Carried into the 29 Oct forecast as 25 Nov through the agent."',
 'O59':'Weeks 2 to 4 are worked back from Xero total cash less the other accounts ($456,746, the gap on 5 Oct). Replace each with the CommBiz cheque + savings figure at 14, 21 and 28 Oct. Week 1 is the same method off 7 Oct.',
 'O60':"Should be nil once everything's in. Weeks 2 to 4 show the gap between the Xero P&L method and the bank - card spend, transfers and anything not on a bill. The CommBiz balances will tighten it.",
}
for k,v in notes.items(): cf[k]=v
cf['A69']="1. Aware Super paid the $225,887 on 13 Oct and everything else by 27 Oct - risk cleared. The whole Complete AV bill ($178,119) went out ahead of it. Build-up on the Complete AV tab."
cf['A72']="4. BAS - Sep quarter (~$96k) did not go out on 28 Oct. Confirm whether it is lodged through the agent (25 Nov) - it is in the 29 Oct forecast on that date. Swap the GST estimate for the activity statement figure."
wb.calculation.fullCalcOnLoad=True
wb.save('out/oct_actuals.xlsx'); print('ok')
