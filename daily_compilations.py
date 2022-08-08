
import pandas as pd
import numpy as np
from tqdm.notebook import tqdm_notebook as tqdm
tqdm.pandas()
from tqdm import tqdm
import os
import getopt,sys
import daily_config as config


def get_nan_inds(series):
    ''' Obtain the first and last index of each consecutive NaN group.
    '''
    series = series.reset_index(drop=True)
    index = series[series.notna()].index.to_numpy()
    if len(index) == 0:
        return []
    indices = np.split(index, np.where(np.diff(index) > 1)[0] + 1)
     
    groups=[(ind[0], ind[-1]+1) for ind in indices]
    return groups

def get_drill_data(data,indicators,comp_target=''):
    data=data=data.T.reset_index(drop=True).T

    data_columns=[]
    ind_head_groups=[]
    for col in data.columns:
        blocks=get_nan_inds(data[col])
        #print(blocks)
        block_list=[]
        for rows in blocks:
            start=rows[0]
            end=rows[-1]
            if start==end:
                end+=1
            data_col=data.loc[list(range(start,end)),col]
            block_list.append(data_col)
            for thing in indicators:
                if thing in data_col.to_list():
                    size=rows[1]-rows[0]
                    #print((thing,rows,size))
                    ind_head_groups.append((thing,size))

        try:
            non_nan=pd.concat(block_list,axis=0)
            data_columns.append(non_nan)
        except Exception as e:
            #print(e)
            pass
    data=pd.concat(data_columns,axis=1,ignore_index=True).reset_index(drop=True)
    data=data.astype('object')

    headers=[]
    for i,thing in enumerate(indicators):
        header=data.where(data.eq(thing)).stack().index.values[0][0]
        if thing in ind_head_groups[i]:
            size=ind_head_groups[i][1]
            #print(thing,'header:',header,'size:',size)
        headers.append((header,size))
    headers
    slices=[]
    for i,header in enumerate(headers):
        head=header[0]
        size=header[1]
        columns = data.iloc[head].str.lower()
        
        try:
            start=head+1
            end=start+size
            
            #print(head,'-',end)
            slice = pd.DataFrame(data.values[start:end], columns=columns).reset_index(drop=True)
        except Exception as e: print('Exception:',e)
        slices.append(slice)
    clean_data=pd.concat(slices,axis=1)

    clean_data=clean_data.dropna(how='all',axis=1).T.dropna(how='all',axis=0).T

    clean_data=clean_data.T.groupby(level=0).first().T
    if comp_target==''or comp_target=='water':
        for col in clean_data.columns:
            try:clean_data[col]=clean_data[col].str.replace(' ','')
            except Exception as e: pass
            try:    clean_data[col]=clean_data[col].str.replace('**','',regex=False)
            except Exception as e: pass
            try:    clean_data[col]=clean_data[col].str.replace('!','',regex=False)
            except Exception as e: pass
            try:    clean_data[col]=clean_data[col].str.strip('ft')
            except Exception as e: pass
            try:    clean_data[col]=clean_data[col].str.strip('gal')
            except Exception as e: pass
            try:    clean_data[col]=clean_data[col].replace(r'\s',' ',regex=True)
            except Exception as e: pass
            try:clean_data[col]=pd.to_numeric(clean_data[col])
            except Exception as e:pass
        try:    clean_data.columns=[col.replace('\n','') for col in clean_data.columns]
        except Exception as e: print(e)

        clean_data['drill from']=pd.to_numeric(clean_data['drill from'],errors='coerce')
        drop=clean_data[clean_data['drill from'].isna()==True].index
        clean_data=clean_data.drop(drop)
        drop_col=['name','shift','size','tools']
        for col in drop_col:
            try:clean_data=clean_data.drop(col,axis=1).fillna(0)
            except Exception as e: print(e)
        clean_data['total man hours']=pd.to_numeric(clean_data['total man hours'],errors='coerce').fillna(method='ffill').fillna(method='bfill')
        dr=data.where(data.eq('Drill')).stack().index.values[0]
        clean_data['drill']=data.iloc[dr[0],dr[1]+1]
        try:
            clean_data.rename(columns={'volume':'water_gal'},inplace=True)
        except Exception as e: print(e)

    if comp_target=='activities':
        clean_cols=['hole no.']
        for col in clean_cols:
            try:    clean_data[col]=clean_data[col].str.replace('**','',regex=False)
            except Exception as e: pass
            try:    clean_data[col]=clean_data[col].str.replace('!','',regex=False)
            except Exception as e: pass
            try:    clean_data[col]=clean_data[col].replace(r'\s',' ',regex=True)
            except Exception as e: pass
            try:    clean_data[col]=pd.to_numeric(clean_data[col])
            except Exception as e:pass
        try:    clean_data.columns=[col.replace('\n','') for col in clean_data.columns]
        except Exception as e: print(e)
        drop_col=['description','unit','quantity']
        for col in drop_col:
            try:clean_data=clean_data.drop(col,axis=1).fillna(0)
            except Exception as e: print(e)
        try: 
            hole=data.where(data.eq('Hole(s)')).stack().index.values[0]
            hole=data.iloc[hole[0],hole[1]+1]
            #print(hole)
            clean_data['hole no.']=hole
        except Exception as e: print(e)
        drop=clean_data[clean_data['activity']==False].index
        clean_data=clean_data.drop(drop)
    date=data.where(data.eq('Date')).stack().index.values[0]
    clean_data['date']=data.iloc[date[0],date[1]+1]

    return clean_data

def main(argv):
    path=[]
    try:
        opts, args = getopt.getopt(argv,"ri:o:",["input_path=","output_file="])
        for opt, arg in opts:
            if opt == '-r':
                print ('daily_compilations.py -i <input_path> -o <output_path>')
                print('using defaults if no file specified')
                
            elif opt in ("-i", "--input_path"):
                path = arg
                print (f'Input files path is {arg} ',)
            elif opt in ("-o", "--output_path"):
                output_path= arg
                print ('Output file path is ', output_path)
    except getopt.GetoptError as e:
        print (e)
        print (' Path READ ERROR use defaults')
    if len(path)==0:
        path=config.path
        output_path=path
    print ('Input path is ', path)
    print ('Output path is ', output_path)
    name_list=[]
    data_list=[]

    for folder in os.listdir(path):
        level2=path+folder+'/'
        print(level2)
        if os.path.isdir(level2):
            for i,file in enumerate(os.listdir(level2)):
                print(i,file)
                file_path=level2+file

                try:
                    data=pd.read_excel(file_path,na_filter=True)
                    #data['FileName']=file
                    data_list.append(data)
                    name_list.append(file_path)
                except Exception as e:
                    print(e)


    water_indicators=['Total Man Hours','Hole No.']
    act_indicators=['Description']
    water_data=[]
    act_data=[]
    for data in data_list:
        w_data=get_drill_data(data=data,indicators=water_indicators,comp_target='water')
        a_data=get_drill_data(data=data,indicators=act_indicators,comp_target='activities')

        water_data.append(w_data)
        act_data.append(a_data)
    final_water=pd.concat(water_data,axis=0)
    final_act=pd.concat(act_data,axis=0)

    final_water=final_water.sort_values(['hole no.','date','drill from']).reset_index(drop=True)
    final_cols=['hole no.','date','drill from','drill to','water_gal']
    others=[col for col in final_water.columns if col not in final_cols]
    final_water=pd.concat([final_water[final_cols],final_water[others]],axis=1)
    final_act=final_act.sort_values(['hole no.','date']).reset_index(drop=True)
    print('output',output_path+'activity_compilation.xlsx')
    final_act.to_excel(output_path+'activity_compilation.xlsx',index=False)
    print('output',output_path+'water_compilation.xlsx')
    final_water.to_excel(output_path+'water_compilation.xlsx',index=False)


if __name__ == "__main__":
    main(sys.argv[1:])


