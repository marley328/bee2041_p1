import openpyxl, pandas as pd, datetime
F='/root/.claude/uploads/d82c9cc5-7263-5271-8116-79a86939a805/fbf94fb6-VA_Revenue_EPS_Diluted_Operating_Actuals_Consensus_Beat_Miss_2025_2029_AsOf_2026-11-18.xlsx'
def sheet(wb,n):
    rows=list(wb[n].iter_rows(values_only=True)); return pd.DataFrame(rows[1:],columns=rows[0])
def load():
    wb=openpyxl.load_workbook(F,data_only=True,read_only=True)
    ctl={r[0]:r[1] for r in wb['Control'].iter_rows(values_only=True) if r[0]}
    a=sheet(wb,'Actuals'); c=sheet(wb,'Consensus'); li=sheet(wb,'Line_Items')
    a=a[a.Period_Type=='Annual']; c=c[(c.Period_Type=='Annual')&(c.Lookback_Days==0)]
    num=lambda s: pd.to_numeric(s,errors='coerce')
    a['val']=num(a.Actual_Value); c['val']=num(c.Consensus_Value)
    a['reported']=pd.to_numeric(a.Value_Date,errors='coerce').notna()
    m=a.merge(c[['Ticker','Line_Item','Period_Label','val']],on=['Ticker','Line_Item','Period_Label'],suffixes=('_a','_c'))
    m['value']=m.val_a.where(m.reported, m.val_c)
    m['status']=m.reported.map({True:'A',False:'E'})
    return m, ctl, a
if __name__=='__main__':
    m,ctl,a=load()
    print('snapshot',ctl.get('Snapshot date'))
    print(a[a.Actual_Value.apply(lambda x:isinstance(x,str))][['Ticker','Line_Item','Period_Label','Actual_Value']])
    m['cell']=m.value.round(2).astype(str)+m.status
    print(m.pivot_table(index=['Ticker','Line_Item'],columns='Fiscal_Year',values='cell',aggfunc='first').to_string())
    print('diff actual vs cons(0) for unreported:', (m[~m.reported].val_a-m[~m.reported].val_c).abs().max())
