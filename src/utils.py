
def getKeyEnglish(str):
    str = str.strip()
    if str == "公司地址":
        return "address"
    elif str == "地址":
        return "address"
    elif str == "姓名":
        return "name"
    elif str == "电话":
        return "mobile"
    elif str == "网址":
        return "url"
    elif str == "职务":
        return "job"
    elif str == "公司英文名称":
        return "companyname"
    elif str == "公司英文名称2":
        return "companyname2"   
    elif str == "标语":
        return "slogan"
    elif str == "公司名称":
        return "company"   
    elif str == "公司名称2":
        return "company2"
    elif str == "邮箱":
        return "email"   
    elif str == "Logo":
        return "logo"   
    elif str == "Logo2":
        return "logo2"
    elif str == "二维码":
        return "qrcode"   
    elif str == "QQ":
        return "qq"   
    elif str == "公众号":
        return "qbjnews"   
    elif str == "固定电话":
        return "phone"   
    else:
        return ""