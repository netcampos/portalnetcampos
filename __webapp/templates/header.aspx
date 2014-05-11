<!DOCTYPE html>
<html>
<head>
 <%=_app.getHeader()%>   
</head>
<body>
<div id="fb-root"></div> <div id="gnc-twitter"></div> <div id="gnc-googleplus"></div>

<!-- start: Fixed Menu -->
<header id="FixedTopBar">
	<div class="bg">
        <div class="container">
        	<% if _app.userON() then %>
            	
            <% else %>
                <div class="pull-left">
                    <ul>
                        <li class="icoHome"></li>
                        <li class="TotalMembers"><strong><%=_app.totalMembers()%></strong> pessoas já estão conectadas ao <strong>Portal NetCampos</strong>, participe você também.</li>
                    </ul>                    
                </div>
                <div class="pull-right">
                    <ul>
                        <li><a href="http://camposdojordao.netcampos.com/?redir=http://www.netcampos.com/#entrar" rel="nofollow" class="gncToolTip" data-toggle="tooltip" data-placement="bottom" title="entre agora" ><i class="icon icon-off"></i><span class="text">Entrar</span></a></li>
                        <li><a href="http://camposdojordao.netcampos.com/cadastro/" class="gncToolTip" data-toggle="tooltip" data-placement="bottom" title="faça seu cadastro"><i class="icon icon-user"></i><span class="text">Cadastre-se</span></a></li>
                        <li><a href="<%=_lib.domainName()%>/contato/" rel="nofollow" class="gncToolTip" data-toggle="tooltip" data-placement="bottom" title="fale conosco"><i class="icon icon-envelope"></i><span class="text">Contato</span></a></li>
                    </ul>
                </div>            
            <% end if %>
        </div>
    </div>
</header>
<!-- end: Fixed Menu -->