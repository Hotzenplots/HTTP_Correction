import concurrent.futures

def Swimming_Pool(Para_Functional_Function,Para_Some_Iterable_Obj):
    with concurrent.futures.ThreadPoolExecutor(max_workers=35) as Pool_Executor:
        Pool_Executor.map(Para_Functional_Function,(Para_Some_Iterable_Obj))