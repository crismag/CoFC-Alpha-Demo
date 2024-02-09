
import sys, os
import requests
import xml.etree.ElementTree as ET
import base64

script_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(script_dir)

class SoapClient:
    def __init__(self, **kwargs):
        self.wsdl_url = kwargs.get('wsdl_url', None)
        self.username = kwargs.get('username', None)
        self.password = kwargs.get('password', None)
        self.proxy_url = kwargs.get('proxy_url', None)
        self.proxy_username = kwargs.get('proxy_username', None)
        self.proxy_password = kwargs.get('proxy_password', None)
        self.session = requests.Session()
        self.downloaded_file = None
        self.headers = {
            'Content-Type': 'text/xml;charset=utf-8',
            'SOAPAction': ''
        }

        if self.username and self.password:
            auth_string = f'{self.username}:{self.password}'
            encoded_auth_string = base64.b64encode(auth_string.encode()).decode()
            self.headers['Authorization'] = f'Basic {encoded_auth_string}'
        if self.proxy_url:
            proxies = {'http': self.proxy_url, 'https': self.proxy_url}
            self.session.proxies.update(proxies)

            if self.proxy_username and self.proxy_password:
                proxy_auth = requests.auth.HTTPProxyAuth(self.proxy_username, self.proxy_password)
                self.session.auth = proxy_auth

    def build_request(self, soap_action, soap_body):
        soap_envelope = '''
        <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:get="http://www.example.com/plm/getChangeXML">
            <soapenv:Body>
                {0}
            </soapenv:Body>
        </soapenv:Envelope>
        '''.format(soap_body)
        self.headers['SOAPAction'] = soap_action
        return soap_envelope

    def send_request(self, soap_action, soap_body):
        soap_request = self.build_request(soap_action, soap_body)
        try:
            response = self.session.post(self.wsdl_url, data=soap_request, headers=self.headers)
            response.raise_for_status()
            soap_response = response.content
            return soap_response
        except requests.exceptions.RequestException as e:
            print("Error during request", str(e))
        except ET.ParseError as e:
            print("Error parsing SOAP response", str(e))
        except Exception as e:
            print("General Error ", str(e))

    def extract_value(self, soap_response):
        '''
        :param soap_response:
            <S:Envelope xmlns:S="http://schemas.xmlsoap.org/soap/envelope/">
                <S:Body>
                    <ns0:getFilesResponse xmlns:ns0="http://www.example.com/plm/getChangeXML">
                    <return>
                        <fileNames>MR-008-XX04_34.docx</fileNames>
                        <files>
                            <fileContent>CONTENT_IN_BASE_64</fileContent>
                            <fileName>MR-008-XX034_34.docx</fileName>
                        </files>
                        <message>success</message>
                        <status>success</status>
                    </return>
                </ns0:getFilesResponse>
                </S:Body>
            </S:Envelope>
            <S:Envelope xmlns:S="http://schemas.xmlsoap.org/soap/envelope/">

            SAMPLE FAILED MESSAGE:
            <S:Envelope xmlns:S="http://schemas.xmlsoap.org/soap/envelope/">
               <S:Body>
                <ns0:getFilesResponse xmlns:ns0="http://www.example.com/plm/getChangeXML">
                <return>
                    <fileNames>MR-008-XX034_34.docx</fileNames>
                    <message>MR-008-XX032_34.docx cannot be find in MR-008-XX034</message>
                    <status>fail</status>
                </return>
                </ns0:getFilesResponse>
               </S:Body>
            </S:Envelope>
        :param xpath:
        :return:
        '''
        namespaces = {
            'S': 'http://schemas.xmlsoap.org/soap/envelope/',
            'nso': 'http://www.example.com/plm/getChangeXML'
        }
        root = ET.fromstring(soap_response)
        status = root.find(".//status", namespaces).text
        message = root.find(".//message", namespaces).text
        print("Request Status:")
        print("Status    : ", status)
        print("Message   : ", message)
        try:
            fileNames = root.find(".//fileNames", namespaces).text
            print("FileNames : ", fileNames)
        except Exception as e:
            print("FileNames : <>")

        if status.lower() == 'success':
            file_elements = root.findall(".//files", namespaces)
            for node in file_elements:
                fn = node.find(".//fileName" , namespaces).text
                file_content_base64 = node.find(".//fileContent").text
                file_content = base64.b64decode(file_content_base64)
                with open(fn, 'wb') as file:
                    file.write(file_content)
                    print(f"Saved file : {fn}")
                    self.downloaded_file = fn
        print("Soap Query Completed.")

    def plm_build_request(self, **kwargs):
        docNum = kwargs.get('docNum', None)
        ecnNum = kwargs.get('ecnNum', None)
        ftypes = kwargs.get('ftypes', None)
        fnames = kwargs.get('fnames', None)
        soap_args = []
        if docNum:
            soap_args.append("<arg0>" + str(docNum) + "</arg0>")
        if ecnNum:
            soap_args.append("<arg1>" + str(ecnNum) + "</arg1>")
        if ftypes:
            soap_args.append("<arg2>" + str(ftypes) + "</arg2>")
        if fnames:
            soap_args.append("<arg3>" + str(fnames) + "</arg3>")
        if soap_args:
            bodyint = ''.join(soap_args)
        else:
            return
        soap_body = '<get:getFiles>{0}</get:getFiles>'.format(bodyint)
        return soap_body

