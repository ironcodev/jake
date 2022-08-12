<%
    dim IS_DEVELOPMENT, IS_PRODUCTION, DEBUG_MODE, DEBUG_USERS

    DEBUG_MODE = true
    DEBUG_USERS = ""

    IS_DEVELOPMENT = true
    IS_PRODUCTION = Not IS_DEVELOPMENT

    dim DB_SERVER, DB_NAME, DB_USER, DB_PASS, DB_PATH, DB_CON_STR

    If IS_DEVELOPMENT Then
        DB_SERVER = "."
        DB_NAME = "MyDb"
        DB_USER = "user1"
        DB_PASS = "12345"

        'DB_PATH = Server.MapPath("../../data/mydb.mdb")
        'DB_CON_STR = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & DB_PATH
        'DB_CON_STR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_PATH
        'DB_CON_STR = "Provider=SQLNCLI;Server=.;Database=" & DB_NAME & ";Uid=" & DB_USER & ";Pwd=" & DB_PASS & ";"
        'DB_CON_STR = "Provider=MSOLEDBSQL;Server=.;Database=" & DB_NAME & ";UID=" & DB_USER & ";PWD=" & DB_PASS & ";"
    Else
        DB_SERVER = "."
        DB_NAME = "MyDb"
        DB_USER = "user1"
        DB_PASS = "12345"
    End If
    
    DB_CON_STR = "Provider=SQLNCLI11;Server=" & DB_SERVER & ";Database=" & DB_NAME & ";Uid=" & DB_USER & ";Pwd=" & DB_PASS & ";"
    
    public const SOCIAL_NETWORK_GITHUB = "https://github.com/mycompany/myproduct"

    public const COMPANY_TITLE = "My Company"
    public const COMPANY_NAME = "MyCompany"
    public const COMPANY_URL = "https://mycompany.com"
    
    dim SITE_ADDR

    If IS_DEVELOPMENT Then
        SITE_ADDR = "http://localhost"
    Else
        SITE_ADDR = "http://www.myproduct.com"
    End If

    public const PRODUCT_TITLE = "My Product"
    public const PRODUCT_NAME = "MyProduct"
    public const PRODUCT_DESC = "My Product offers a new kind of product to users"
    public const PRODUCT_URL = "https://www.myproduct.com"

    public const PRODUCT_PAGE_TITLE = "My Product - A new kind of product"

    public const MAIL_DEFAULT = "info@myproduct.com"
    public const MAIL_FROM = "noreply@myproduct.com"
    public const MAIL_SERVER = "mail.myproduct.com"
    public const MAIL_SSL = False
    public const MAIL_PORT = 587
    public const MAIL_AUTHENTICATE = True
    public const MAIL_TIMEOUT = 25
    public const MAIL_USER = "noreply@myproduct.com"
    public const MAIL_PASS = "12345"

    dim MAIL_TEMPLATE_FORGOTPASSWORD

    MAIL_TEMPLATE_FORGOTPASSWORD = _
        "Hi<br/>" & _
        "<p>This email is sent to you from <b>" & PRODUCT_TITLE & "</b> to reset your password.</p>" & _
        "<p>Please click the link below or copy/paste its address into your browser to reset your password.</p>" & _
        "<p>" & _
        "<center>" & _
        "<b><a href=""" & SITE_ADDR & "/resetpassword/?token={token}"">Reset My Password</a></b>" & _
        "<br/>" & _
        "<br/>" & _
        "<a href=""" & SITE_ADDR & "/resetpassword/?token={token}"">Reset My Password</a>" & _
        "</center>" & _
        "</p>" & _
        "<p>If you have not issued a password reset request you can safely delete this message.</p>" & _
        "<p>Kind Regards</p>" & _
        "<p>" & PRODUCT_NAME & " Support</p>"

    public const CRYPT_KEY = "123abc456"
    public const HASH_CSS = "100"
    public const HASH_JS = "100"

    dim SITE_LANG

    SITE_LANG = "en"
%>