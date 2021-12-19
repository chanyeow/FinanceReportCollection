import openpyxl
import math
import os
import time
import requests
import xlrd
import platform
import os.path
import sys
import xml.dom.minidom

from pathlib import Path
from xml.dom.minidom import parse

def writeXMLtest(isUpdate, filename):
    file = Path(filename)
    if not file.is_file():
        doc = xml.dom.minidom.Document() 
        with open(filename, 'w') as f:
            doc.writexml(f, addindent='  ', encoding='utf-8')
        
    domTree = parse(filename)
    
    # 文档根元素
    rootNode = domTree.documentElement

    # 新建一个customer节点
    customer_node = domTree.createElement("customer")
    customer_node.setAttribute("ID", "C003")

    # 创建name节点,并设置textValue
    name_node = domTree.createElement("name")
    name_text_value = domTree.createTextNode("kavin")
    name_node.appendChild(name_text_value)  # 把文本节点挂到name_node节点
    customer_node.appendChild(name_node)

    # 创建phone节点,并设置textValue
    phone_node = domTree.createElement("phone")
    phone_text_value = domTree.createTextNode("32467")
    phone_node.appendChild(phone_text_value)  # 把文本节点挂到name_node节点
    customer_node.appendChild(phone_node)

    # 创建comments节点,这里是CDATA
    comments_node = domTree.createElement("comments")
    cdata_text_value = domTree.createCDATASection("A small but healthy company.")
    comments_node.appendChild(cdata_text_value)
    customer_node.appendChild(comments_node)

    rootNode.appendChild(customer_node)

    with open(filename, 'w') as f:
        # 缩进 - 换行 - 编码
        domTree.writexml(f, addindent='  ', encoding='utf-8')

def test():
    #在内存中创建一个空的文档
    doc = xml.dom.minidom.Document() 
    #创建一个根节点Managers对象
    root = doc.createElement('Managers') 
    #设置根节点的属性
    root.setAttribute('company', 'xx科技') 
    root.setAttribute('address', '科技软件园') 
    #将根节点添加到文档对象中
    doc.appendChild(root)

    managerList = [{'name' : 'joy', 'age' : 27, 'sex' : '女'},
    {'name' : 'tom', 'age' : 30, 'sex' : '男'},
    {'name' : 'ruby', 'age' : 29, 'sex' : '女'}
    ]

    for i in managerList :
        nodeManager = doc.createElement('Manager')
        nodeName = doc.createElement('name')
        #给叶子节点name设置一个文本节点，用于显示文本内容
        nodeName.appendChild(doc.createTextNode(str(i['name'])))

        nodeAge = doc.createElement("age")
        nodeAge.appendChild(doc.createTextNode(str(i["age"])))

        nodeSex = doc.createElement("sex")
        nodeSex.appendChild(doc.createTextNode(str(i["sex"])))

        #将各叶子节点添加到父节点Manager中，
        #最后将Manager添加到根节点Managers中
        nodeManager.appendChild(nodeName)
        nodeManager.appendChild(nodeAge)
        nodeManager.appendChild(nodeSex)
        root.appendChild(nodeManager)

    #开始写xml文档
    fp = open('Manager.xml', 'w')
    doc.writexml(fp, indent='\t', addindent='\t', newl='\n', encoding="utf-8")

if __name__ == '__main__':
    test()