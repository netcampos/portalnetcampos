Imports Microsoft.VisualBasic
Imports System
Imports System.Net
Imports System.XML

public class _weather

	Public function agora(codigo as string) as System.XML.XmlNodeList
							
		Dim xml_weather as XmlDocument
		Dim url_weather as String = "http://xoap.weather.com/weather/local/" & codigo & "?cc=*&dayf=5&link=xoap&prod=xoap&par=1094583197&key=d74181c7e17dbea3&unit=m"
		Dim reader_weather as XmlTextReader = new XmlTextReader(url_weather)
		reader_weather.Read()
		
		If reader_weather.HasValue Then	
			reader_weather.WhitespaceHandling = WhitespaceHandling.Significant
			xml_weather = new XmlDocument()
			Dim nodes As XmlNodeList
			xml_weather.Load(reader_weather)
			nodes = xml_weather.SelectNodes("item")
			try
				return nodes
			catch ex as exception
				
			end try
		end if
	
	End Function
	
	Public Function hoje(ByVal codigo as string) as System.XML.XmlNodeList
	
		Dim xml_weather as XmlDocument
		Dim url_weather as String = "http://xoap.weather.com/weather/local/" & codigo & "?cc=*&dayf=5&link=xoap&prod=xoap&par=1094583197&key=d74181c7e17dbea3&unit=m"
		Dim reader_weather as XmlTextReader = new XmlTextReader(url_weather)
		reader_weather.Read()
		
		If reader_weather.HasValue Then	
			reader_weather.WhitespaceHandling = WhitespaceHandling.Significant
			xml_weather = new XmlDocument()
			Dim nodes As XmlNodeList
			xml_weather.Load(reader_weather)
			nodes = xml_weather.SelectNodes("weather")
			try
				return nodes
			catch ex as exception
				
			end try
		end if	
	
	End Function
	
	Public Function hojeCptec(ByVal codigo as string) as System.XML.XmlNodeList
	
		Dim xml_weather as XmlDocument
		Dim url_weather as String = "http://servicos.cptec.inpe.br/XML/cidade/1219/previsao.xml"
		try
			xml_weather = new XmlDocument()
			Dim nodes As XmlNodeList
			xml_weather.Load(url_weather)
			nodes = xml_weather.SelectNodes("cidade")
			return nodes
		catch ex as exception
			
		end try
	
	End Function	
	

end class