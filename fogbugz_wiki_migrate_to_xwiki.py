'''
fogbugs_downloader.py :: uses fogbugz api to download wikis with attachments to xwiki format. Uses pycookiecheat
to download attachments (see here https://stackoverflow.com/q/51498991/471160).

NOTE1: attachments download is not very friendly, I was able to make it work only on macos, I had to be logedin
      to fogbugz on chrome to make the attachment retrieval work.

NOTE2: this script uses modifiee fogbugz api (https://developers.fogbugz.com/default.asp?W199). The only modification
        is to allow add http simple auth credentials.
'''
import re
import shutil
import sys
import time
import html
import os
import traceback
import ssl
import xml.etree.ElementTree as ET
import base64
from io import BytesIO

import yaml
from bs4 import BeautifulSoup

# For win32 pycookiecheat tweaks see: https://github.com/luskan/pycookiecheat
import urllib

from urllib3.exceptions import InsecureRequestWarning
from xml.dom import minidom

from pycookiecheat import chrome_cookies

import requests

from fogbugz_v1 import FogBugz

userhome = os.path.expanduser('~')

if sys.platform == 'darwin':
    cookies_path = userhome + '/Library/Application Support/Google/Chrome/Default/Cookies'
elif sys.platform.startswith('win32'):
    #cookies_path = userhome + r'\AppData\Local\Google\Chrome\User Data\Default\Cookies'
    #cookies_path = userhome + r'\AppData\Local\Google\Chrome\User Data\Profile 1\Cookies';
    cookies_path = userhome + r'\AppData\Local\Google\Chrome\User Data\Profile 2\Cookies'
else:
    cookies_path = userhome + r'/.config/google-chrome/Default/Cookies'

def dump_to_xwiki(config):
    """
    Generates xwiki xar (xwiki.org)
    """

    fogbugz_url = config['url']
    attachment_destination_url = config['attachment_destination_url']
    api_version = config['api']
    login_email = config['email']
    login_password = config['password']
    token = config['token']
    auth_user = config.get('auth_user', None)
    auth_password = config.get('auth_password', None)
    wiki_dir_name_parent = config.get('wiki_dir_name_parent', None)
    wiki_dir_name = config['wiki_dir_name']
    wiki_root_name_dir = wiki_dir_name_parent + os.path.sep + wiki_dir_name + os.path.sep + config['wiki_root_name']
    wiki_root_name_ref = wiki_dir_name.replace('.', '\\.') + '.' + config['wiki_root_name'].replace('.', '\\.')
    wiki_root_name = wiki_dir_name + '.' + config['wiki_root_name'].replace('.', '\\.')
    wiki_root_name_pure = config['wiki_root_name']

    print("Dumping wiki for: " + fogbugz_url)

    # (monkey patching) Allow for self signed ssl certs
    ssl._create_default_https_context = ssl._create_unverified_context

    #
    fb = FogBugz(fogbugz_url, token,
                 api_version=api_version,
                 auth_user=auth_user,
                 auth_password=auth_password)
    fb.logon(login_email, login_password)

    # Index all wikis, this allows for further links fixing
    respWikis = fb.listWikis()
    article_to_wiki_dir = {}
    for wiki in respWikis.find_all('wiki'):
        respArticles = fb.listArticles(ixWiki=wiki.ixWiki.string)
        for article in respArticles.find_all('article'):
            if article.ixWikiPage.string in article_to_wiki_dir:
                raise Exception("Wiki article id already exists! : " + article.ixWikiPage.string + " in " + article_to_wiki_dir[article.ixWikiPage.string])
            article_to_wiki_dir[article.ixWikiPage.string] = wiki_dir_name + os.path.sep + config['wiki_root_name'] + os.path.sep + wiki.sWiki.string + os.path.sep + article.ixWikiPage.string

    #
    if not os.path.isdir("./" + wiki_dir_name):
        os.makedirs("./" + wiki_dir_name)

    package_xml_path = "./" + wiki_dir_name_parent + os.sep + "package.xml"
    if os.path.isfile(package_xml_path):
        package = ET.parse(package_xml_path).getroot()
    else:
        package = ET.Element('package')

    package_infos = package.find('infos')
    if package_infos is None:
        package_infos = ET.SubElement(package, 'infos')
    package_infos = package

    package_infos_name = package_infos.find('name')
    if package_infos_name is None:
        package_infos_name = ET.SubElement(package_infos, 'name')
    package_infos_name.text = wiki_dir_name.replace('.', '\\.') + ".WebHome"

    package_infos_description = package_infos.find('description')
    if package_infos_description is None:
        package_infos_description = ET.SubElement(package_infos, 'description')

    package_infos_licence = package_infos.find('licence')
    if package_infos_licence is None:
        package_infos_licence = ET.SubElement(package_infos, 'licence')

    package_infos_author = package_infos.find('author')
    if package_infos_author is None:
        package_infos_author = ET.SubElement(package_infos, 'author')
    package_infos_author.text = ''

    package_infos_version = package_infos.find('version')
    if package_infos_version is None:
        package_infos_version = ET.SubElement(package_infos, 'version')

    package_infos_backupPack = package_infos.find('backupPack')
    if package_infos_backupPack is None:
        package_infos_backupPack = ET.SubElement(package_infos, 'backupPack')
    package_infos_backupPack.text = "false"

    package_infos_preserveVersion = package_infos.find('preserveVersion')
    if package_infos_preserveVersion is None:
        package_infos_preserveVersion = ET.SubElement(package_infos, 'preserveVersion')
    package_infos_preserveVersion.text = "false"

    package_files = package.find('files')
    if package_files is None:
        package_files = ET.SubElement(package, 'files')

    # This is required to make import work
    package_files = package

    rootWebHomeText = wiki_dir_name.replace('.', '\\.') + ".WebHome"
    rootWebHomeFile = [el for el in package.findall('.//file') if el.text == rootWebHomeText]
    if not rootWebHomeFile:
        package_file = ET.SubElement(package_files, 'file', defaultAction='0', language='0')
        package_file.text = rootWebHomeText

    package_file = ET.SubElement(package_files, 'file', defaultAction='0', language='0')
    package_file.text = wiki_root_name_ref + ".WebHome"

    os.makedirs("." + os.path.sep + wiki_root_name_dir)
    myfile = open("." + os.path.sep + wiki_root_name_dir + os.path.sep + "WebHome.xml", "w")
    wiki_base = ET.Element('xwikidoc')
    wiki_base.set("version", "1.3")
    wiki_base.set("reference", wiki_root_name_ref + ".WebHome")
    wiki_base.set("locale", "")
    ET.SubElement(wiki_base, 'web').text = wiki_root_name
    ET.SubElement(wiki_base, 'name').text = "WebHome"
    ET.SubElement(wiki_base, 'language')
    ET.SubElement(wiki_base, 'defaultLanguage').text = "en"
    ET.SubElement(wiki_base, 'translation').text = "0"
    ET.SubElement(wiki_base, 'creator').text = "xwiki:XWiki.nemo"
    ET.SubElement(wiki_base, 'creationDate').text = "1588827562000"
    ET.SubElement(wiki_base, 'parent').text = wiki_dir_name.replace('.', '\\.') + ".WebHome"
    ET.SubElement(wiki_base, 'author').text = "xwiki:XWiki.nemo"
    ET.SubElement(wiki_base, 'contentAuthor').text = "xwiki:XWiki.nemo"
    ET.SubElement(wiki_base, 'date').text = "1588827562000"
    ET.SubElement(wiki_base, 'contentUpdateDate').text = "1588827562000"
    ET.SubElement(wiki_base, 'version').text = "1.1"
    ET.SubElement(wiki_base, 'title').text = wiki_root_name_pure
    ET.SubElement(wiki_base, 'comment').text = "Imported from XAR"
    ET.SubElement(wiki_base, 'minorEdit').text = "false"
    ET.SubElement(wiki_base, 'syntaxId').text = "xwiki/2.1"
    ET.SubElement(wiki_base, 'hidden').text = "false"
    ET.SubElement(wiki_base, 'content').text = "{{children/}}"
    wiki_base_xml = ET.tostring(wiki_base, encoding='unicode')
    wiki_base_xml_reparsed = minidom.parseString(wiki_base_xml)
    wiki_base_xml = wiki_base_xml_reparsed.toprettyxml(indent="\t")
    myfile.write(wiki_base_xml)

    for wiki in respWikis.find_all('wiki'):
        package_file = ET.SubElement(package_files, 'file', defaultAction='0', language='0')
        package_file.text = wiki_root_name_ref + "." + wiki.sWiki.string + ".WebHome"

        current_wiki_dir = "." + os.path.sep + wiki_root_name_dir + os.path.sep + wiki.sWiki.string
        os.mkdir(current_wiki_dir)

        myfile = open(current_wiki_dir + "/WebHome.xml", "w")
        wiki_base = ET.Element('xwikidoc')
        wiki_base.set("version", "1.3")
        wiki_base.set("reference", wiki_root_name_ref + "." + wiki.sWiki.string + ".WebHome")
        wiki_base.set("locale", "")
        ET.SubElement(wiki_base, 'web').text = wiki_root_name + "." + wiki.sWiki.string
        ET.SubElement(wiki_base, 'name').text = "WebHome"
        ET.SubElement(wiki_base, 'language')
        ET.SubElement(wiki_base, 'defaultLanguage').text = "en"
        ET.SubElement(wiki_base, 'translation').text = "0"
        ET.SubElement(wiki_base, 'creator').text = "xwiki:XWiki.nemo"
        ET.SubElement(wiki_base, 'creationDate').text = "1588827562000"
        ET.SubElement(wiki_base, 'parent').text = wiki_root_name_ref + ".WebHome"
        ET.SubElement(wiki_base, 'author').text = "xwiki:XWiki.nemo"
        ET.SubElement(wiki_base, 'contentAuthor').text = "xwiki:XWiki.nemo"
        ET.SubElement(wiki_base, 'date').text = "1588827562000"
        ET.SubElement(wiki_base, 'contentUpdateDate').text = "1588827562000"
        ET.SubElement(wiki_base, 'version').text = "1.1"
        ET.SubElement(wiki_base, 'title').text = wiki.sWiki.string
        ET.SubElement(wiki_base, 'comment').text = "Imported from XAR"
        ET.SubElement(wiki_base, 'minorEdit').text = "false"
        ET.SubElement(wiki_base, 'syntaxId').text = "xwiki/2.1"
        ET.SubElement(wiki_base, 'hidden').text = "false"
        ET.SubElement(wiki_base, 'content').text = "{{children/}}"
        wiki_base_xml = ET.tostring(wiki_base, encoding='unicode')
        wiki_base_xml_reparsed = minidom.parseString(wiki_base_xml)
        wiki_base_xml = wiki_base_xml_reparsed.toprettyxml(indent="\t")
        myfile.write(wiki_base_xml)

        respArticles = fb.listArticles(ixWiki=wiki.ixWiki.string)
        for article in respArticles.find_all('article'):
            try:
                full_article = fb.viewArticle(ixWikiPage=article.ixWikiPage.string)
            except fb.FogBugzConnectionError as err:
                print("fogbugz_v1.FogBugzConnectionError: for wiki={0}, message={1}"
                      .format(err, article.ixWikiPage.string))
                return

            wiki_page_id = article.ixWikiPage.string
            os.mkdir(current_wiki_dir + "/" + wiki_page_id)

            #package_file = ET.SubElement(package_files, 'file', defaultAction='0', language='0')
            #package_file.text = wiki_root_name_ref + "." + wiki.sWiki.string + "." + wiki_page_id + ".WebHome"

            data = ET.Element('xwikidoc')
            data.set("version", "1.3")
            data.set("reference", wiki_root_name_ref + "." + wiki.sWiki.string + "." + wiki_page_id + ".WebHome")
            data.set("locale", "")

            ET.SubElement(data, 'web').text = wiki_root_name + "." + wiki.sWiki.string + "." + wiki_page_id
            ET.SubElement(data, 'name').text = "WebHome"
            ET.SubElement(data, 'language')
            ET.SubElement(data, 'defaultLanguage').text = "en"
            ET.SubElement(data, 'translation').text = "0"
            ET.SubElement(data, 'creator').text = "xwiki:XWiki.nemo"
            ET.SubElement(data, 'creationDate').text = "1588827562000"
            ET.SubElement(data, 'parent').text = wiki_root_name_ref + "." + wiki.sWiki.string+".WebHome"
            ET.SubElement(data, 'author').text = "xwiki:XWiki.nemo"
            ET.SubElement(data, 'contentAuthor').text = "xwiki:XWiki.nemo"
            ET.SubElement(data, 'date').text = "1588827562000"
            ET.SubElement(data, 'contentUpdateDate').text = "1588827562000"
            ET.SubElement(data, 'version').text = "1.1"
            ET.SubElement(data, 'title').text = str(article.sHeadline.string)
            ET.SubElement(data, 'comment').text = "Imported from XAR"
            ET.SubElement(data, 'minorEdit').text = "false"
            ET.SubElement(data, 'syntaxId').text = "xwiki/2.1"
            ET.SubElement(data, 'hidden').text = "false"
            content = full_article.sBody.string
            if content is None:
                content = ""

            soup = BeautifulSoup(content, "html.parser")
            for tag in soup.find_all(['a', 'img']):
                if tag.has_attr('href') or tag.has_attr('src'):
                    m = None
                    if tag.has_attr('href'):
                        m = re.match(r'default\.asp\?W(\d+)', tag['href'], re.M|re.I)
                    if m:
                        new_url = attachment_destination_url + \
                                    urllib.parse.quote(article_to_wiki_dir[m.group(1)])

                        linkName = '[[' + tag.text + '>>' + new_url + ']]'
                        linkName = '[[' + new_url + ']]'
                        urlLink = soup.new_string(linkName)
                        tag.replace_with(urlLink)
                    else:
                        img_width = -1
                        img_height = -1
                        if tag.has_attr('href'):
                            m = re.match(r'default\.asp\?pg=pgDownload.*sFileName=([^;]+)', tag['href'], re.M | re.I)
                        if not m and tag.has_attr('src'):
                            m = re.match(r'default\.asp\?pg=pgDownload.*sFileName=([^;]+)', tag['src'], re.M | re.I)
                            if m and tag.has_attr('width'):
                                img_width = int(tag['width'])
                            if m and tag.has_attr('height'):
                                img_height = int(tag['height'])
                        if m:
                            if tag.name == 'a':
                                download_url = tag['href'].replace('amp;', '')
                            elif tag.name == 'img':
                                download_url = tag['src'].replace('amp;', '')
                            new_url = fogbugz_url + html.unescape(download_url)

                            att_file_name = m.group(1)
                            print(" --- %s:%s " % (att_file_name, new_url))

                            cookies = chrome_cookies(new_url, cookie_file=cookies_path)

                            if tag.name == 'a':
                                linkName = '[[attach:' + att_file_name + '||target="_blank"]]'
                            elif tag.name == 'img':
                                #[[image:img.png||width="25" height="25"]]
                                linkName = '[[image:' + att_file_name + '||' +\
                                           ('width="' + str(img_width) + '" ' if img_width != -1 else '') +\
                                           ('height="' + str(img_height) + '"' if img_height != -1 else '') +\
                                           ']]'

                            attachLink = soup.new_string(linkName)
                            tag.replace_with(attachLink)

                            encodedAtt = ""
                            sizeAtt = 0
                            for tries in range(10):
                                try:
                                    session = requests.Session()
                                    if auth_user and auth_password:
                                        session.auth = (auth_user, auth_password)
                                    requests.packages.urllib3.disable_warnings(category=InsecureRequestWarning)
                                    with session.get(new_url, cookies=cookies, stream=True, verify=False) as r:
                                        r.raise_for_status()
                                        if r.status_code == 200:
                                            att_path = os.path.join(current_wiki_dir, wiki_page_id, att_file_name)
                                            #with open(att_path, 'wb') as f:
                                            f = BytesIO()
                                            for chunk in r.iter_content(chunk_size=8192):
                                                # If you have chunk encoded response uncomment if
                                                # and set chunk_size parameter to None.
                                                # if chunk:
                                                f.write(chunk)
                                            sizeAtt = f.getbuffer().nbytes
                                            encodedAtt = base64.b64encode(f.getvalue()).decode()
                                    break
                                except BaseException as error:
                                    print('An exception at try {} occurred: {}, {}'.format(tries, error, traceback.format_exc()))
                                    time.sleep(tries * 10)

                            attachment = ET.SubElement(data, 'attachment')
                            ET.SubElement(attachment, 'author').text = "xwiki:XWiki.nemo"
                            ET.SubElement(attachment, 'date').text = "1559758114000"
                            ET.SubElement(attachment, 'version').text = "1.1"
                            ET.SubElement(attachment, 'comment').text = ""
                            ET.SubElement(attachment, 'filename').text = att_file_name
                            ET.SubElement(attachment, 'filesize').text = str(sizeAtt)
                            ET.SubElement(attachment, 'content').text = encodedAtt
                        elif tag.name == 'a' and tag.has_attr('href') and tag.has_attr('rel'):
                            # xwiki for some reasons badly formats a-tags, it shows internal tag attributes
                            # so we strip a-hrefs and insert instead only href
                            # rel="nofollow"
                            linkName = '[[' + tag['href'] + ']]'
                            xwikiLink = soup.new_string(linkName)
                            tag.replace_with(xwikiLink)

            content = str(soup)
            #article_to_wiki

            ET.SubElement(data, 'content').text = "{{html wiki=\"true\"}}" + content + "{{/html}}"

            mydata = ET.tostring(data, encoding='unicode')
            reparsed = minidom.parseString(mydata)
            mydata = reparsed.toprettyxml(indent="\t")

            myfile = open(current_wiki_dir + "/" + wiki_page_id + "/" + "WebHome.xml", "w")
            myfile.write(mydata)

            print(wiki_page_id)


    #
    # Save WebHome.xml for root dir
    myfile = open("." + os.sep + wiki_dir_name_parent + os.path.sep + wiki_dir_name + os.sep + "WebHome.xml", "w")
    wiki_base = ET.Element('xwikidoc')
    wiki_base.set("version", "1.3")
    wiki_base.set("reference", wiki_dir_name.replace('.', '\\.') + ".WebHome")
    wiki_base.set("locale", "")
    ET.SubElement(wiki_base, 'web').text = wiki_dir_name
    ET.SubElement(wiki_base, 'name').text = "WebHome"
    ET.SubElement(wiki_base, 'language')
    ET.SubElement(wiki_base, 'defaultLanguage').text = "en"
    ET.SubElement(wiki_base, 'translation').text = "0"
    ET.SubElement(wiki_base, 'creator').text = "xwiki:XWiki.nemo"
    ET.SubElement(wiki_base, 'creationDate').text = "1588827562000"
    ET.SubElement(wiki_base, 'parent').text = "common.WebHome" #"internal:Main2.WebHome"
    ET.SubElement(wiki_base, 'author').text = "xwiki:XWiki.nemo"
    ET.SubElement(wiki_base, 'contentAuthor').text = "xwiki:XWiki.nemo"
    ET.SubElement(wiki_base, 'date').text = "1588827562000"
    ET.SubElement(wiki_base, 'contentUpdateDate').text = "1588827562000"
    ET.SubElement(wiki_base, 'version').text = "1.1"
    ET.SubElement(wiki_base, 'title').text = wiki_dir_name
    ET.SubElement(wiki_base, 'comment').text = "Imported from XAR"
    ET.SubElement(wiki_base, 'minorEdit').text = "false"
    ET.SubElement(wiki_base, 'syntaxId').text = "xwiki/2.1"
    ET.SubElement(wiki_base, 'hidden').text = "false"
    ET.SubElement(wiki_base, 'content').text = "{{children/}}"
    wiki_base_xml = ET.tostring(wiki_base, encoding='unicode')
    wiki_base_xml_reparsed = minidom.parseString(wiki_base_xml)
    wiki_base_xml = wiki_base_xml_reparsed.toprettyxml(indent="\t")
    myfile.write(wiki_base_xml)

    #
    # Save package.xml
    mydata = prettify(package)
    #mydata = ET.tostring(package, encoding='unicode')
    #reparsed = minidom.parseString(mydata)
    #mydata = reparsed.toprettyxml(indent="\t")
    myfile = open("./"+wiki_dir_name_parent + os.sep + "package.xml", "w")
    myfile.write(mydata)


def prettify(root):
    for elem in root.iter('*'):
        if elem.text is not None:
            elem.text = elem.text.strip()
        if elem.tail is not None:
            elem.tail = elem.tail.strip()

    rough_string = ET.tostring(root, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="\t")

def main():
    configs = yaml.safe_load(open("./settings.yml"))

    if configs.get("dir_to_rm", None):
        if os.path.isdir("./" + configs['dir_to_rm']):
            try:
                shutil.rmtree("./" + configs['dir_to_rm'])
            except:
                print("Error deleting root dir!")
                pass
            if os.path.isdir("./" + configs['dir_to_rm']):
                print("Error deleting root dir!")
                pass

    for config in configs['servers']:
        dump_to_xwiki(config)

if __name__ == '__main__':
    main()
