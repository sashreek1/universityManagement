import json
from pyexcel_ods import get_data

def get_data_academics (odspath, final_dict):
    for k in range(1,5):
        sheets = get_data(odspath+"/Data"+str(k)+'.ods')
        sheets = json.loads((json.dumps(sheets)))
        data_dict = {}
        sub_list = ["", "Algebra", "Physics", "PE", "Chemistry", "Geometry", "Biology", "Programming"]
        lst= []
        for sheet in sheets:
            lst = sheets[sheet][3:]
        while [] in lst: lst.remove([])
        for i in range(0,len(lst)):
            person = None
            for j in range(len(lst[i])):
                if isinstance(lst[i][j],str):
                    data_dict[lst[i][j]] = {}
                    person = lst[i][j]
                else:
                    data_dict[person][sub_list[j]] = lst[i][j]
            else:
                data_dict[person]['Math'] = (data_dict[person]['Geometry']+data_dict[person]['Algebra'])/2
        for person in data_dict:
            if k ==1:
                final_dict[person] = {}
            average = ((2*data_dict[person]["Physics"]) + data_dict[person]["Chemistry"] + 2*(data_dict[person]["Math"]) + data_dict[person]["Programming"] + data_dict[person]["Biology"] + data_dict[person]["PE"])/8
            final_dict[person]['academics'+str(k)] = average


def get_data_ielts(odspath, final_dict):
    sheets = get_data(odspath+"/IELTS.ods")
    sheets = json.loads((json.dumps(sheets)))
    data_dict = {}
    lst= []
    for sheet in sheets:
        lst = sheets[sheet][3:]
    for i in range(len(lst)):
        temp_dict = {}
        person = lst[i][0]
        score = sum(lst[i][1:])/4
        temp_dict['IELTS'] = score
        data_dict[person] = temp_dict
        final_dict[person].update(temp_dict)

def get_data_interview(odspath, final_dict):
    sheets = get_data(odspath+"/Interview.ods")
    sheets = json.loads((json.dumps(sheets)))
    data_dict = {}
    lst= []
    for sheet in sheets:
        lst = sheets[sheet][3:]
    for i in range(len(lst)):
        temp_dict = {}
        person = lst[i][0]
        score = sum(lst[i][1:])/5
        temp_dict['interview'] = score
        data_dict[person] = temp_dict
        final_dict[person].update(temp_dict)


def scan_ods (data_folder):
    dic = {}
    get_data_academics(data_folder,dic)
    get_data_ielts(data_folder,dic)
    get_data_interview (data_folder,dic)
    return dic