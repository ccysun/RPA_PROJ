import time, Parameters
import UserModule
from ast import literal_eval

def Wait(term=0):
    """해당 시간만큼 대기하는 함수

    Activity: COMMON

    Common:

    API:
        - term은 0이상의 값이어야 하나 음수가 전달된 경우 대기없이 통과함(0초 대기)
        - 주의 : term의 값은 milliseconds로 전달해야 함
    Activity:


    :param term: 대기 시간
    :return: None
    """
    print('MY Debug', Parameters.value)
    if "term" in Parameters.value:
        try:
            term += int(Parameters.value["term"])
        except:
            pass

    # term이 음수로 넘어온 경우
    if term < 0:
        term = 0

    time.sleep(term / 1000)

def baEval(in_string):
    try:
        obj = literal_eval(in_string)
    #except ValueError:
    except:
        try:
            corrected = "\'" + in_string + "\'"
            obj = literal_eval(corrected)
        except:
            obj = in_string
    return obj

def CustomActivity(func_name="", arg1="", arg2="", arg3="", arg4="", arg5=""):
    
    func = Parameters.value['func_name']
    arg1 = Parameters.value['arg1']
    arg2 = Parameters.value['arg2']
    arg3 = Parameters.value['arg3']
    arg4 = Parameters.value['arg4']
    arg5 = Parameters.value['arg5']
        
    UserModule.register_functions()
        
    if(arg1 == ""):
        result = UserModule.custom_func(func, None)
        
    elif(arg2 == ""):
        result = UserModule.custom_func(func, baEval(arg1))
        
    elif(arg3 == ""):
        result = UserModule.custom_func(func, baEval(arg1), baEval(arg2))
        
    elif(arg4 == ""):
        result = UserModule.custom_func(func, baEval(arg1), baEval(arg2), baEval(arg3))
        
    elif(arg5 == ""):
        result = UserModule.custom_func(func, baEval(arg1), baEval(arg2), baEval(arg3), baEval(arg4))
        
    else:
        result = UserModule.custom_func(func, baEval(arg1), baEval(arg2), baEval(arg3), baEval(arg4), baEval(arg5))
    
    return result