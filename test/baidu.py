import urllib.request
import urllib.parse
import json
import hashlib
import time, re

encoding = 'utf-8' # 编码用utf-8
salt = '666' #随机数
appid = '20200211000382472'
secret_key = 'iRE4E5_Ebf2QzZj19tS3'
# 请求失败码
REQUEST_FAILED = -1

# 正则匹配
settings_regex = r"\s*\'.+\'\s*=>\s*.+"


def getMD5(content):
    m2 = hashlib.md5()
    m2.update(content.encode(encoding))
    return m2.hexdigest()


def getTranslateResponce(url, data):
    data = urllib.parse.urlencode(data).encode('utf-8')
    response = urllib.request.urlopen(url, data)
    return response.read().decode('utf-8')


def trans(content):
    url = 'http://api.fanyi.baidu.com/api/trans/vip/translate'
    data = {}
    data['appid'] = appid
    data['salt'] = salt
    data['from'] = 'auto'
    data['to'] = 'en'
    data['q'] = content
    data['sign'] = getMD5(appid + content + salt + secret_key)
    html = getTranslateResponce(url, data)
    target = json.loads(html)
    while target.get('error_code', REQUEST_FAILED) != REQUEST_FAILED:
        # print('本次请求失败，原因为：',target['error_msg'])
        time.sleep(1)
        html = getTranslateResponce(url, data)
        target = json.loads(html)
    # print(target)
    return target['trans_result'][0]['dst']


def translate():
    file = open('en.php', 'r')
    output = open('cn.php', 'w')
    for line in file.readlines():
        # print(line)
        if re.match(settings_regex, line):
            result = re.search(r"\'[^=]+\'", line)
            # original_text 待翻译文本
            original_text = result.group()
            translated_text = trans(original_text).lower()
            translated_text = translated_text.replace('”', '')
            translated_text = translated_text.replace('“', '')
            if translated_text[0] != "'":
                translated_text = "'" + translated_text
            if translated_text[-1] != "'":
                translated_text = translated_text + "'"
            # print(translated_text)
            # print(original_text,' => ',translated_text)
            line = "    " + original_text + " => " + translated_text + ",\n"
            line = line.replace('"', "'")
        output.write(line)
        output.flush()
        print("写入：" + line, end='')
    file.close()
    output.close()


if __name__ == '__main__':
    # translate()
    text = trans("把图片3往左边移动10厘米,再往右移一点点,还要放大点")
    print(text)



# 中文：这是一张桌子，我要去上学了，你慢点走"
# 百度翻译 This is a table. I'm going to school. Slow down  
# 微软翻译 It's a table, I'm going to school, you go slow


中文："把图片3往左边移动10厘米,再往右移一点点,还要放大点"
百度：Move picture 3 10 cm to the left, a little bit to the right, and a little bit larger
微软：Move Picture 3 to the left by 10 cm, then to the right a little bit, and zoom in