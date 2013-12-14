package commons;

import java.util.Random;

public class Tool {

	
	/**
	 * 分页
	 * @param page
	 * @return
	 */
	public static String  getPage(Integer page){
		StringBuffer sb = new StringBuffer();
			sb.append("<div id='page' class='paginationPage'>");
			sb.append("<div class='pagination-item'>");
		if(page==null || page <=5){
			if(page==null)
					sb.append("<a target='_self' href='#' id='page_1' onclick='toPage(1)'  class='current'>1</a>");
			else{
				for(int i=1;i<10;i++){
					if(page==i)
						sb.append("<a target='_self' href='#' id='page_"+i+"' onclick='toPage("+i+")'  class='current'>"+i+"</a>");
					else
						sb.append("<a target='_self' href='#' id='page_"+i+"' onclick='toPage("+i+")'>"+i+"</a>");
				}
			}
				
		}else{
			sb.append("<a target='_self' id='page_1' onclick='toPage(1)' href='#'>首页</a> ");
			for(int i=page-3;i<page;i++){
				sb.append("<a target='_self' id='page_"+i+"' onclick='toPage("+i+")' href='#'>"+i+"</a> ");
			}
				sb.append("<a target='_self' href='#' id='page_"+page+"' onclick='toPage("+page+")'  class='current'>"+page+"</a>");
			for(int i=page+1;i<page+2;i++){
				sb.append("<a target='_self' id='page_"+i+"' onclick='toPage("+i+")' href='#'>"+i+"</a> ");
			}
			    sb.append("<a target='_self' onclick='toPage("+(page+5)+")'  href='#'>...</a> ");
			for(int i=page+7;i<page+10;i++){
				sb.append("<a target='_self' id='page_"+i+"' onclick='toPage("+i+")' href='#'>"+i+"</a> ");
			}    
		}
		    sb.append("</div>");
		    sb.append("<a target='_self' href='#' onclick='toPage("+(page+1)+")'  class='pagination-next' title='下一页'>下一页</a>");
		    sb.append("</div>");
		return sb.toString();
	}
	
	
 	 /**
 	  * 获取数字跟字母的随机16位。 返回大写。    
 	  * @param length   需要输出多少位
 	  * @return
 	  */
 	 public static String getKey(Integer length){
 		String val = "";   
	      Random random = new Random();   
	      for(int i = 0; i < length; i++)   
	      {   
	          String charOrNum = random.nextInt(2) % 2 == 0 ? "char" : "num"; // 输出字母还是数字   
	                 
	          if("char".equalsIgnoreCase(charOrNum)) // 字符串   
	          {   
	              int choice = random.nextInt(2) % 2 == 0 ? 65 : 97; //取得大写字母还是小写字母   
	              val += (char) (choice + random.nextInt(26));   
	          }   
	          else if("num".equalsIgnoreCase(charOrNum)) // 数字   
	          {   
	              val += String.valueOf(random.nextInt(10));   
	          }   
	      }   
	     String vall=val.toUpperCase();
	     vall=vall.replaceAll("0", "A");
	     vall=vall.replaceAll("0", "A");
	     vall=vall.replaceAll("O", "X");
	     vall=vall.replaceAll("o", "X");
	     return vall;
 	 }
 	 
   	/**
   	 *  把数字转换为定长字符串，不足的前补0， 如Num2S(12,6) 返回：000012  如Num2S(99,4) 返回：0099
   	 *  @param m 数字 
   	 *  @param len  目标字符串长度
   	 *  @return 补零后的定长字符串
   	 */
   	
 	public static String Num2S(long m,int len){
 		String s =(m+"").replaceAll("\\.", "");
 	    int c=s.length();
 	    for(int i=0;i<len-c;i++){
 	    	s ="0"+s;
 	    }
 	    return s;
   	}
 	
 	
 	
	/**
	 * 分页
	 * @param page
	 * @return
	 */
	public static String  getBootstartPage(Integer page){
		StringBuffer sb = new StringBuffer();
			sb.append("<div class='pagination'><ul>");
			sb.append("<li><a href='#'onclick='toPage("+(page==null||page==1?1:page-1)+")')>上一页</a></li>");
		if(page==null || page <=5){
			for(int i=1;i<10;i++){
				if(page==i)
					sb.append("<li><a href='#' id='page_"+i+"' class='btn-success' onclick='toPage("+i+")'>"+i+"</a></li>");
				else
					sb.append("<li><a href='#' id='page_"+i+"' onclick='toPage("+i+")'>"+i+"</a></li>");
			}
		}else{
			for(int i=page-3;i<page;i++){
				sb.append("<li><a href='#' id='page_"+i+"' onclick='toPage("+i+")'>"+i+"</a></li>");
			}
			sb.append("<li><a href='#' id='page_"+page+"' class='btn-success' onclick='toPage("+page+")'>"+page+"</a></li>");
			
			for(int i=page+1;i<page+5;i++){
				sb.append("<li><a href='#' id='page_"+i+"' onclick='toPage("+i+")'>"+i+"</a></li>");
			}
		}
			sb.append("<li><a href='#'onclick='toPage("+(page+1)+")'>下一页</a></li>");
		    sb.append("</ul></div>");
		return sb.toString();
	}
}
