package com.httpdownload;

import java.io.BufferedReader;
import java.io.ByteArrayOutputStream;
import java.io.Closeable;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Iterator;
import java.util.List;
import java.util.Random;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.net.ssl.SSLContext;
import javax.net.ssl.TrustManager;
import javax.net.ssl.X509TrustManager;

import org.apache.http.Consts;
import org.apache.http.Header;
import org.apache.http.HttpException;
import org.apache.http.HttpHost;
import org.apache.http.HttpRequest;
import org.apache.http.HttpRequestInterceptor;
import org.apache.http.HttpResponse;
import org.apache.http.NameValuePair;
import org.apache.http.RequestLine;
import org.apache.http.client.CookieStore;
import org.apache.http.client.HttpClient;
import org.apache.http.client.entity.UrlEncodedFormEntity;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.client.methods.HttpRequestBase;
import org.apache.http.client.protocol.ClientContext;
import org.apache.http.conn.params.ConnRoutePNames;
import org.apache.http.conn.scheme.Scheme;
import org.apache.http.conn.ssl.SSLSocketFactory;
import org.apache.http.impl.client.BasicCookieStore;
import org.apache.http.impl.client.DefaultHttpClient;
import org.apache.http.message.BasicHeader;
import org.apache.http.message.BasicNameValuePair;
import org.apache.http.protocol.BasicHttpContext;
import org.apache.http.protocol.HttpContext;

import com.google.common.collect.Multimap;



public class HttpClientUtils {

	
	public static HttpClient getHttpClient()throws Exception{
		DefaultHttpClient httpclient = new DefaultHttpClient();
		httpclient.addRequestInterceptor(new HttpRequestInterceptor() {
			public void process(final HttpRequest request,
					final HttpContext context) throws HttpException,
					IOException {
				if (!request.containsHeader("Accept")) {
					request.addHeader("Accept", "*/*");
				}
				if (request.containsHeader("User-Agent")) {
					request.removeHeaders("User-Agent");
				}
				if (request.containsHeader("Connection")) {
					request.removeHeaders("Connection");
				}
				if (request.containsHeader("Host")) {
					request.removeHeaders("Host");
				}
				request.addHeader("Host", "proactive-assessments-itg.corp.hp.com");
				request.addHeader("User-Agent","Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36");
				request.addHeader("Connection", "keep-alive");
			}
		});	
		
		return httpclient;
	}
	public static HttpClient getHttpsClient()throws Exception{
		DefaultHttpClient httpclient = new DefaultHttpClient();
		TrustManager easyTrustManager = new X509TrustManager() {
		   public void checkClientTrusted(java.security.cert.X509Certificate[]x509Certificates, String s) throws java.security.cert.CertificateException {
		    }
		   public void checkServerTrusted(java.security.cert.X509Certificate[]x509Certificates, String s) throws java.security.cert.CertificateException {
		    }
		   public java.security.cert.X509Certificate[] getAcceptedIssuers() {
		       return null;
		    }
		};
		SSLContext sslcontext =SSLContext.getInstance("TLS");
		sslcontext.init(null, new TrustManager[]{easyTrustManager}, null);
		SSLSocketFactory sf = new SSLSocketFactory(sslcontext,SSLSocketFactory.ALLOW_ALL_HOSTNAME_VERIFIER);
		Scheme sch = new Scheme("https",443,sf);
		httpclient.getConnectionManager().getSchemeRegistry().register(sch);
	
		httpclient.addRequestInterceptor(new HttpRequestInterceptor() {
			public void process(final HttpRequest request,
					final HttpContext context) throws HttpException,
					IOException {
				if (!request.containsHeader("Accept")) {
					request.addHeader("Accept", "*/*");
				}
				if (request.containsHeader("User-Agent")) {
					request.removeHeaders("User-Agent");
				}
				if (request.containsHeader("Connection")) {
					request.removeHeaders("Connection");
				}
				if (request.containsHeader("Host")) {
					request.removeHeaders("Host");
				}
				
				request.addHeader("Host", "proactive-assessments-itg1.corp.hp.com");
				request.addHeader("User-Agent","Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36");
				request.addHeader("Connection", "keep-alive");
			}
		});
		//HttpHost proxy = new HttpHost("web-proxy.atlanta.hp.com", 8088, "http");
		//httpclient.getParams().setParameter(ConnRoutePNames.DEFAULT_PROXY, proxy);
		return httpclient;
	}
	public static HttpResponse sendPost(String url,Multimap<String, String> dataMap,Header... headers)throws Exception{
		HttpPost post = new HttpPost(url);
		return sendRequest(post,dataMap,headers);
	}
	public static HttpResponse sendGet(String url,Header...headers)throws Exception{
		HttpGet get = new HttpGet(url);
		return sendRequest(get,null,headers);
	}
	//support the duplicate key 
	public static HttpResponse sendRequest(HttpRequestBase request,Multimap<String, String> dataMap,Header...headers)throws Exception{
		RequestLine requestLine = request.getRequestLine();
		HttpClient httpclient = null;
		if(requestLine.getUri().toLowerCase().startsWith("https")){
			System.out.println("Request uri:"+requestLine.getUri().toLowerCase());
			httpclient = getHttpsClient();
		}else{
			httpclient = getHttpClient();
		}
		
		//HttpHost proxy = new HttpHost("web-proxy.atlanta.hp.com", 8088, "http");
		//httpclient.getParams().setParameter(ConnRoutePNames.DEFAULT_PROXY, proxy);
		
		
		// the data form
		if(dataMap != null && dataMap.size() > 0){
			List <NameValuePair> nvps = new ArrayList <NameValuePair>();
			for(String key: dataMap.keySet()){
				Collection<String> indexkey=dataMap.get(key);
				if(indexkey!=null){
					for (Iterator<String> iterator = indexkey.iterator(); iterator.hasNext();){
						String value=iterator.next();
					    nvps.add(new BasicNameValuePair(key,value ));
					 }
				}else{
					 nvps.add(new BasicNameValuePair(key,""));
				}
	           
			}
			System.out.println("the request's header we had configured is :"+nvps);
			((HttpPost)request).setEntity(new UrlEncodedFormEntity(nvps, Consts.UTF_8));
			//remove it if existing
			if (request.containsHeader("Content-Length")) {
				request.removeHeaders("Content-Length");
			}
			if (request.containsHeader("Content-Type")) {
				request.removeHeaders("Content-Type");
			}
			//request.setHeader("Content-Length",""+((HttpPost)request).getEntity().getContentLength());
			request.setHeader("Content-Type","application/x-www-form-urlencoded");
			
		}else{
			 System.out.println("the request data parameter head is empty");
		}
		if(headers != null && headers.length > 0){
			request.setHeaders(headers);
		}
		//System.out.println("execute request now...");
		CookieStore cookieStore = new BasicCookieStore();
		HttpContext httpContext = new BasicHttpContext();
		httpContext.setAttribute(ClientContext.COOKIE_STORE, cookieStore);
		return  httpclient.execute(request,httpContext);
	}
	
	public static Header[] getRequestCookieHeader(HttpResponse response){
		Header[] responseHeaders = response.getHeaders("Set-Cookie");
		if(responseHeaders == null || responseHeaders.length <= 0){
			System.out.println("the request set cookies is null");
			return null;
		}
		Header[] requestCookies = new BasicHeader[responseHeaders.length];
		for(int i=0;i<responseHeaders.length;i++){
			 requestCookies[i] = new BasicHeader("Cookie",responseHeaders[i].getValue());
		}
		return requestCookies;
	}
	
	public static Header[] getHttpsCookie(String httpsIndexUrl)throws Exception{
		
		HttpResponse response = sendGet(httpsIndexUrl,(Header[])null);
		Header[] httpsCookieHeaders = getRequestCookieHeader(response);
		
		System.out.println("get the http cookie:"+java.util.Arrays.toString(httpsCookieHeaders));
		return httpsCookieHeaders;
	}
	
	public static Header[] buildRequestHeader(String referurl,Header[] cookieHeaders){
		
		Header referHeader = new BasicHeader("Referer",referurl);
		Header[] headers = new Header[1+cookieHeaders.length];
		headers[0] = referHeader;
		for(int i=0;i<cookieHeaders.length;i++){
			headers[i+1] = cookieHeaders[i];
		}
		return headers;
	}
	
	public static String getStrutsToken(String content){
		if(content == null || content.equals("")){
			return "";
		}
		Matcher m = Pattern.compile("(?is)<input .*?name=\"org.apache.struts.taglib.html.TOKEN\".*?value=\"(\\w+)\".*/?>").matcher(content);
		if(m.find()){
			System.out.println("Find the struts2 token is:"+m.group(1));
			return m.group(1);
		}
		return "";
	}
	public static String getLeftTicketToken(String content){
		if(content == null || content.equals("")){
			return "";
		}
		Matcher m = Pattern.compile("(?is)<input.*?id=\"left_ticket\".*?value=\"(\\w+)\".*/?>").matcher(content);
		if(m.find()){
			System.out.println("get the left ticket token is:"+m.group(1));
			return m.group(1);
		}
		return "";
	}
	public static String getSeatAvailable(String content){
		if(content == null || content.equals("")){
			return "";
		}
		StringBuilder sb = new StringBuilder();
		Matcher m = Pattern.compile("(?is)<select.*?id=\"passenger_1_seat\".*?</select>").matcher(content);
		if(m.find()){
			String select = m.group();
			//System.out.println("seat avaliable:"+select);
			m = Pattern.compile("(?is)<option.*?>(.*?)</option>").matcher(select);
			while(m.find()){
				sb.append(m.group(1)).append(",");
			}
			if(sb.length() > 0){
				sb.deleteCharAt(sb.length()-1);
			}
		}
		return sb.toString();
	}
	public static String getContent(HttpResponse response)throws Exception{
		InputStream input=response.getEntity().getContent();
		StringBuilder sb = new StringBuilder();
		BufferedReader br = new BufferedReader(new InputStreamReader(input));
		String line = "";
		while ((line = br.readLine()) != null) {
			sb.append(line);
			sb.append("\n");
		}
		close(input,br);
		return sb.toString();
	}
	public static byte[] readInputStream(InputStream in)throws Exception{
		byte[] buff = new byte[1024];
		ByteArrayOutputStream bo = new ByteArrayOutputStream();
		int len = 0;
		while((len = in.read(buff))>-1){
			bo.write(buff,0,len);
		}
		close(bo,in);
		return bo.toByteArray();
	}
	public static boolean saveInputStreamAsPic(InputStream in)throws Exception{
		Random rand=new Random();
		int rnd =rand.nextInt(900 - 1) + 1;
		FileOutputStream out = new FileOutputStream("C:\\passcode\\passCodeAction"+rnd+".tif");
		byte[] buff = new byte[1024];
		int len = 0;
		while((len = in.read(buff)) > -1){
			//System.out.println("图片大小："+len);
			out.write(buff,0,len);
		}
		close(in,out);
		return true;
	}
	
	public static void close(Closeable... closeables){
		if(closeables != null){
			for(Closeable closeable : closeables){
				try{
					closeable.close();
				}catch(Exception e){
					e.printStackTrace();
				}
			}
		}
	}
	
}
