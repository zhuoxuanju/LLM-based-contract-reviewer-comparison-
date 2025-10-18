# -*- coding: utf-8 -*-
from Class.API import llm_review_result
import codecs
import hashlib
import os
import platform
import subprocess
from datetime import datetime
import io
import docx
from flask import Flask, request, abort, jsonify, Response
from werkzeug.exceptions import BadRequest
from flask_jwt_extended import create_access_token, jwt_required, JWTManager, get_jwt_identity
from flask_login import login_required
from Class.Contract import Contract
from docx import Document


from threading import Thread
from threading import Event

# 防止乱码
# sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='gb18030')

app = Flask(__name__)
app.config['UPLOAD_EXTENSIONS'] = ['docx', 'doc']
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32 MB
app.config['JWT_SECRET_KEY'] = 'super-secret'
jwt = JWTManager(app)

filename_cache = ""
file_name = ""


# @app.errorhandler(Exception)
# def handle_error(e):
#     exc_type = type(e).__name__
#     exc_msg = str(e)
#     response = {
#         'status': 'error',
#         'message': f'{exc_type}: {exc_msg}'
#     }
#     return jsonify(response), 500


@app.before_request
def log_each_request():
    app.logger.info('【请求方法】{}【请求路径】{}【请求地址】{}'.format(
        request.method, request.path, request.remote_addr))


@app.route('/')
def hello_world():  # put application's code here
    return 'Hello World!'


def write_to_file(message):
    with codecs.open('a.txt', 'a', encoding='utf-8') as file:
        file.write(message + '\n')


@app.route('/postfile', methods=['POST'])
# @jwt_required()
def postfile():
    file = request.files['file']
    filename = file.filename
    global file_name
    file_name = filename
    if filename.split('.')[-1] not in app.config['UPLOAD_EXTENSIONS']:
        # return HTTP 400 error for invalid file extensions
        abort(400, description='Invalid file extension')
    if file.content_length > app.config['MAX_CONTENT_LENGTH']:
        # return HTTP 413 error for large files
        abort(413, description='File too large')
    suffix = filename.split('.')[-1]
    filename = "tmp_" + datetime.now().strftime('%Y-%m-%d_%H-%M-%S.')
    folder_name = "files"
    file_path = os.path.join(folder_name, filename + suffix)
    os.makedirs(folder_name, exist_ok=True)
    file.save(file_path)
    try:
        if suffix == "doc":
            sys = platform.system()
            if sys == 'Windows':
                import pythoncom
                import win32com.client
                dir_path = os.path.dirname(
                    os.path.abspath(__file__)) + '\\files\\'
                pythoncom.CoInitialize()
                word = win32com.client.Dispatch('Word.Application')
                doc = word.Documents.Open(dir_path + filename + "doc")
                doc.SaveAs(dir_path + filename + 'docx', 12)
                doc.Close()
                pythoncom.CoUninitialize()
                os.remove(dir_path + filename + suffix)
            else:
                current_dir = os.getcwd()

                # 构建输入和输出文件的相对路径
                input_file = os.path.join(
                    current_dir, "files/" + filename + 'doc')
                output_file = os.path.join(
                    current_dir, "files/" + filename + 'docx')
                print(input_file, output_file)
                # 使用 unoconv 将 DOC 文件转换为 DOCX 格式
                subprocess.call(
                    ['unoconv', '-f', 'docx', '-o', output_file, input_file])
                os.remove(input_file)
        global filename_cache
        filename_cache = filename + 'docx'

        #####################################
        # 后台执行新审查
        #####################################
        llm_review_result.start(review_background)

        return {'file_path': '/files/' + filename_cache, 'file_name': file_name, "code": 200}
    except BadRequest:
        # return HTTP 400 error for invalid files
        abort(401, description='Convert fail')


@app.route('/getfile')
# @jwt_required()
def getfilepath():
    return {'file_path': '/files/' + filename_cache, 'file_name': file_name, "code": 200}


@app.route('/download_docx')
# @jwt_required()
def download_docx():
    # 从文件中读取.docx文件
    docx_path = 'files/' + filename_cache
    try:
        doc = Document(docx_path)
    except:
        return Response(None)

    # 创建一个内存缓冲区来保存文件流
    output = io.BytesIO()

    # 将.docx文件写入内存缓冲区
    doc.save(output)

    # 将内存缓冲区的指针移至文件的开头
    output.seek(0)

    # 获取内存缓冲区的内容
    file_content = output.getvalue()

    # 返回文件内容作为响应体给前端
    return Response(file_content,
                    content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


@app.route('/review', methods=['POST'])
# @jwt_required()
def review():
    try:
        contract = Contract()
        # 若版本号匹配失败，则由前端传入
        if request.args.get('version'):
            contract.version = request.args.get('version')
        document = docx.Document('files/' + filename_cache)
        # modified by wsr 20231018
        # 取合同读取结果，如果get_text返回-1，则通过大模型进行信息抽取
        read_flag = contract.get_text(document)
        # 模板设置成功才进行内容审查
        if read_flag != -1 and contract.rules is not None:
            contract.match_basic_info()
            contract.match_all()
        if read_flag == -1:
            contract.match_with_llm()
        result = contract.get_result()
        result["code"] = 200
        return result
    except BadRequest:
        abort(400, description='Invalid file')


def encode_to_md5(password):
    md5 = hashlib.md5()
    md5.update(password.encode())
    return md5.hexdigest()


@app.route('/login', methods=['POST'])
def login():
    try:
        user_dict = {"dict": "gtH3p9%0JX",
                     "reviewer01": "dict/gtH3p9%0JX",
                     "HeFei01": "dict/3401)!HF",
                     "WuHu02": "dict/3402)@WH",
                     "BengBu03": "dict/3403)#BB",
                     "HuaiNan04": "dict/3404)$HN",
                     "MaAnShan05": "dict/3405)%MAS",
                     "HuaiBei06": "dict/3406)^HB",
                     "TongLing07": "dict/3407)&TL",
                     "AnQing08": "dict/3408)*AQ",
                     "HuangShan09": "dict/3410)(HS",
                     "ChuZhou10": "dict/3411!)CZ",
                     "FuYang11": "dict/3412!@FY",
                     "SuZhou12": "dict/3413!#SZ",
                     "LuAn13": "dict/3424@$LA",
                     "XuanCheng14": "dict/3425@%XC",
                     "ChaoHu15": "dict/3426@^CH",
                     "ChiZhou16": "dict/3429@(CZ"}
        user_name = request.get_json()['username']
        password = request.get_json()['password']
        if user_name not in user_dict.keys():  # 如果用户不存在
            return jsonify({"msg": "未找到用户", "status": 0, "code": 200})

        if encode_to_md5(user_dict[user_name]) != password:  # 如果密码不正确
            return jsonify({"msg": "密码错误", "status": 0, "code": 200})

        access_token = create_access_token(identity=user_name)
        return jsonify({"access_token": access_token, "msg": "登录成功", "status": 1, "code": 200})

    except BadRequest:
        abort(400, description='Invalid user')


def review_background(evnet: Event):
    llm_review_result.do_before_run()

    import sys
    # sys.stdout = open('log/rule.md', 'a', encoding='utf-8')

    contract = Contract()
    document = docx.Document('files/' + filename_cache)
    contract.get_text(document)

    from Class.RuleManager import RULERS_MANAGER
    rulers_manager = RULERS_MANAGER()

    from Class.Searcher import RuleSearcher
    rule_searcher = RuleSearcher(contract)

    from Class.Reviewer import Reviewer
    reveiwer = Reviewer()

    for _, rule in enumerate(list(rulers_manager.get_rulers(rulers_manager.rules_class[0]).rulers.values())[:]):
        if evnet.is_set():
            llm_review_result.stop()
            return False
        review_result = {
            "rule_name": None,
            "search_output": {
                "search_result": None,
                "search_from": None,
            },
            "review_func": None,
            "review_output": {
                "func_result": None,
                "llm_result": None,
            }
        }
        review_result['rule_name'] = rule.rule_name
        try:
            review_result['search_output'] = rule_searcher.search_0205(rule)
        except Exception as e:
            print(e)
        review_result['review_func'] = rule.review_reg_name
        review_result['review_output'] = reveiwer.review(
            rule, review_result['search_output'])
        llm_review_result.append(review_result)
    llm_review_result.end()
    llm_review_result.save("./log/" + file_name + ".json")


@app.route('/review_llm', methods=['POST'])
def review_llm():
    try:
        result = {}
        result["result"] = llm_review_result.to_list
        result["review"] = llm_review_result.STATUS.name
        result["code"] = 200
        return result
    except BadRequest:
        print(BadRequest)
        abort(400, description='Invalid file')


if __name__ == '__main__':
    # app.run()
    app.run(host='127.0.0.1', port=5000)
