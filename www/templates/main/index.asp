<!--#include virtual="/init.asp"-->
<%
	class Template
		public page
		
		public sub render()
%>
<!DOCTYPE html>
<html lang="en"> 
<head>
	<title><%=page.title%></title>
    <!-- Meta -->
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="<%=PRODUCT_DESC%>">
    <meta name="author" content="<%=COMPANY_TITLE%>">
    <link rel="shortcut icon" href="/favicon.ico"> 
    
    <!-- Google Font -->
    <link href="https://fonts.googleapis.com/css?family=Poppins:300,400,500,600,700&display=swap" rel="stylesheet">
    
    <!-- FontAwesome JS-->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.1.2/css/all.min.css" integrity="sha512-1sCRPdkRXhBV2PBLUdRb4tMg1w2YPf37qatUFeS7zlBy7jJI8Lf4VHwWfZZfpXtYSLy85pkm9GaYVYMfw5BC1A==" crossorigin="anonymous" referrerpolicy="no-referrer" />
    
    <!-- Bootstrap icons-->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.5.0/font/bootstrap-icons.css" rel="stylesheet">
    
    <!-- Bootstrap 5.2 -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.2.0/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-gH2yIJqKdNHPEq0n4Mqa/HGKIhSkIHeL5AyhkYV8i59U5AR6csBvApHHNl/vI1Bx" crossorigin="anonymous">
    
    <% page.header() %>

    <!-- Custom CSS -->  
    <link rel="stylesheet" href="/assets/css/styles.css?<%=HASH_CSS%>">
</head> 

<body class="<%=page.bodyClass%>">

	<% page.body() %>
    
	<!--#include file="footer.asp"-->
    
    <!-- Javascript -->          
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.2.0/dist/js/bootstrap.bundle.min.js" integrity="sha384-A3rJD856KowSb7dwlZdYEkO39Gagi7vIsF0jrRAoQmDKKtQBHUuLZ9AsSv4jD4Xa" crossorigin="anonymous"></script>
    <script src="/assets/js/all.min.js?<%=HASH_JS%>"></script> 
    <!-- JavaScript Bundle with Popper -->
    <!-- Page Specific JS -->
	<% page.footer() %>
</body>
</html> 
<%
		end sub
	end class
%>