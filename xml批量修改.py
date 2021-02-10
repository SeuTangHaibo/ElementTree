#coding=utf-8
#__author__=H.Tang

import xml.etree.ElementTree as ET
import pandas as pd
import os 

def indent(elem, level=0):
    i = "\n" + level*"\t"
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "\t"
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
        for elem in elem:
            indent(elem, level+1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i
            
if __name__ == '__main__':
    yc = pd.read_excel(os.path.abspath('.')+'\\天津城东直流储能项目104点表20201121.xlsx',sheet_name='遥测')
    yx = pd.read_excel(os.path.abspath('.')+'\\天津城东直流储能项目104点表20201121.xlsx',sheet_name='遥信')
    tree = ET.parse(os.path.abspath('.')+'\\分布式储能.xml')
    root = tree.getroot()

    for item in root[0].findall('item'):
        root[0].remove(item)
    for item in root[1].findall('item'):
        root[1].remove(item)
    for i in range(len(yc)):
        ET.SubElement(root[0], 'item',attrib = {'entry': str(yc.iloc[i,1]).zfill(3),'desc': str(yc.iloc[i,0])+str(yc.iloc[i,2])})
    for i in range(len(yx)):
        ET.SubElement(root[1], 'item',attrib = {'entry': str(yx.iloc[i,1]).zfill(3),'desc': str(yx.iloc[i,0])+str(yx.iloc[i,2])})
    indent(root)
    tree.write(os.path.abspath('.')+'\\分布式储能1.xml',encoding='utf-8')
