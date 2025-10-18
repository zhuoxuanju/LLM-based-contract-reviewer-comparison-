import codecs
import hashlib
import os
import platform
import subprocess
from datetime import datetime

import docx
from flask import Flask, request, abort
from werkzeug.exceptions import BadRequest
from werkzeug.security import generate_password_hash
import zipfile
from xml.dom.minidom import parse
import xml.dom.minidom
from Class.Contract import Contract

if __name__ == "__main__":
    try:
        # document = docx.Document("testfiles\example_contract.docx")
        document = docx.Document("testfiles\淮北市人民医院新院区机房建设及服务项目技术服务合同.docx")
        # f = zipfile.ZipFile("tmp027.docx", "r")
        #
        # for filename in f.namelist():
        #     f.extract(filename, "extract")  # 解压文件
        #     print(filename)
        #
        # DOMTree = xml.dom.minidom.parse("extract/word/document.xml")
        # collection = DOMTree.documentElement
        # # docs = collection.getElementsByTagName("w:document")
        # print(collection.getAttribute("xmlns:w"))

        #####################################
        # test embedding

        from Class.OPENAI_API import LLM
        # LLM.get_langchian_embedding()
        # exit()

        #####################################

        contract = Contract()
        contract.get_text(document)
        contract.match_basic_info()
        contract.match_all()
        result = contract.get_result()

        # ---------------------
        # ruler test
        # ---------------------
        from Class.RuleManager import RULERS_MANAGER
        rulers_manager = RULERS_MANAGER()

        from Class.Searcher import RuleSearcher
        rule_searcher = RuleSearcher(contract)

        from Class.Reviewer import Reviewer
        reveiwer = Reviewer()

        from Class.OPENAI_API import LLM

        # import sys
        # sys.stdout = open('log/rule.md', 'a', encoding='utf-8')
        

        from Class.Review_func_hub import review_reg_hub
        print(review_reg_hub.regs["is local"]("合肥市庐阳区"))
        # print(review_reg_hub.regs["is company"]("['中电信数智科技有限公司冬季那个分公司', '中电信数智科技有限公司安徽分公司']"))
        exit()

        for i, rule in enumerate(list(rulers_manager.get_rulers(rulers_manager.rules_class[0]).rulers.values())[:2]):
            print("# "+rule.rule_name)
            search_result = rule_searcher.search_0205(rule)
            # rst = reveiwer.review(rule, search_result)
        exit()

        # ---------------------
        # match with Law
        # ---------------------
        import sys
        sys.stdout = open('log/law.md', 'a', encoding='utf-8')
        from Class.logDecorator import print_log

        from Class.Searcher import LawSearcher

        law_search = LawSearcher.from_file("resources\中华人民共和国民法典.txt")

        contract_full = '\n'.join(contract.text)

        import re
        
        for sentence in re.split("第[一二三四五六七八九十]+条|本合同附件为：", contract_full):
            print_log(sentence=sentence)
            output = law_search.detect(sentence)


        # ---------------------
        # talk with doc test
        # ---------------------
        # from Class.TalkWithDoc import TalkWithDoc
        # talk = TalkWithDoc(contract)

        # q_list = [
        #     "合同签订地点在哪里？",
        #     "合同总价是多少，含税价和非含税价分别是多少？",
        # ]
        # for q in q_list:
        #     print(q)
        #     talk.query(q)

        # ---------------------
        # summary test
        # ---------------------        # from Class.Summary_LLM import Summary, prompt_exp
        # print(result['errors'][:])
        # summary = Summary(result['errors'][:])
        # summary.get_summary()
        # print(summary.errors_with_summary)
        # exit()

        '''

        password = "gtH3p9%0JX"
        md5 = hashlib.md5()
        md5.update(password.encode())
        print(md5.hexdigest())
        print(generate_password_hash(password))
        '''
    except BadRequest:
        abort(400, description='Invalid file')
