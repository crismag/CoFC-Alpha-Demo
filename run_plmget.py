import argparse
import os
from CoFCLib import plmget


def read_credentials():
    cred_file_path = os.path.expanduser("~/.cofc_cred/cred")
    credentials = {}
    try:
        with open(cred_file_path, 'r') as cred_file:
            for line in cred_file:
                key, value = line.strip().split('=')
                credentials[key] = value
    except FileNotFoundError:
        print(f"Credential file not found at {cred_file_path}.")
    return credentials


def main():
    parser = argparse.ArgumentParser(description="PLMGET script")
    parser.add_argument('-doc', required=True, help="Specify PLM Document Number.")
    parser.add_argument('-file', required=True, help="Specify the attachment file to download.")
    parser.add_argument('-proxy_user', default=None, help="Specify the proxy username (optional)")
    parser.add_argument('-proxy_pass', default=None, help="Specify the proxy password (optional)")
    args = parser.parse_args()

    # Read credentials from the credential file
    credentials = read_credentials()
    user = credentials.get('user', '')
    password = credentials.get('password', '')

    # Check if proxy_user and proxy_pass are provided, if not, use credentials from the file
    proxy_user = args.proxy_user if args.proxy_user else user
    proxy_pass = args.proxy_pass if args.proxy_pass else password

    print("Downloading files from PLM... Please wait")
    sc = plmget.SoapClient(
        wsdl_url="http://example.com/GetXmlFilesFromChanageWS?wsdl",
        username=user,
        password=password,
        proxy_user=proxy_user,
        proxy_pass=proxy_pass,
    )
    action = 'getFiles'
    plm_request = sc.plm_build_request(
        action=action,
        docNum=args.doc,
        fnames=args.file
    )
    response = sc.send_request(action, plm_request)

    if response is not None:
        sc.extract_value(response)

    print("Download complete.")
    pass


if __name__ == "__main__":
    main()