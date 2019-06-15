'''
Created on 2019年6月14日

@author: Z
'''

import json




if __name__ == '__main__':
    test_dict = {'bigberg': [7600, {1: [['iPhone', 6300], ['Bike', 800], ['shirt', 300]]}]}
    json_str = json.dumps(test_dict)
    '''
    with open("configs.json", "w") as f:
        json.dump(json_str,f)
        print(json_str)
    '''    
    with open("configs.json", "r") as f:
        load_dict = json.load(f)
        
        print(load_dict)
        print(type(load_dict['dddd']))