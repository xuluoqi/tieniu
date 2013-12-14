package controllers;

import controllers.base.WebSiteBaseController;
import play.mvc.*;

import java.util.*;

import commons.Tool;

import models.*;

@With(TopPhotos.class)
public class Application extends WebSiteBaseController {

    public static void index() {
//    	List<ChangePhoto> changes = ChangePhoto.all().fetch();
//    	List<ProductTypes> types = ProductTypes.find("type=null").fetch();
//    	Users user = Users.all().first();
//    	List<Product> products = Product.find("order by id desc").fetch(6);
//        render(changes,types,user,products);
        render2();
    }
    
    
    public static void products(Integer page,Long tid){
    	page = page==null?1:page;
    	Integer bar =2;
    	List<Product> products = new ArrayList<Product>();
    	if(tid==null)
    		products = Product.find("order by id desc").fetch(page,9);
    	else
    		products = Product.find("productType.id=? or productType.type.id=?", tid,tid).fetch(page,9);
    	String pages = Tool.getBootstartPage(page);
    	List<ProductTypes> types = ProductTypes.find("type=null").fetch();
    	for(ProductTypes ts:types){
    		ts.types = ProductTypes.find("type.id=?", ts.id).fetch();
    	}
    	render(bar,products,types,pages,tid);
    }
    
    /**
     * 查看商品详情
     * @param id
     */
    public static void view(Long id){
    	List<ProductTypes> types = ProductTypes.find("type=null").fetch();
    	for(ProductTypes ts:types){
    		ts.types = ProductTypes.find("type.id=?", ts.id).fetch();
    	}
    	Product p = Product.findById(id);
    	Integer bar = 2;
    	render(p,types,bar);
    }
    
    public static void about(){
    	Users user = Users.all().first();
    	List<ProductTypes> types = ProductTypes.find("type=null").fetch();
    	Integer bar = 3;
    	render(bar,user,types);
    }
    
    /**
     * 查看大图
     */
    public static void bigPhoto(){
    	List<Product> products = Product.find("order by id desc").fetch(6);
    	render(products);
    }
    
    /**
     * 联系我们
     */
    public static void messages(){
    	List<ProductTypes> types = ProductTypes.find("type=null").fetch();
    	Integer bar = 4;
    	Users user = Users.all().first();
    	render(user,types,bar);
    }
    
    
	
	public static void login(){
		render();
	}
	
	public static void saveLogin(String email,String password){
		if(email ==null|| password==null){
			flash.put("error","登录账号或密码不能为空!");
			login();
		}else{
			Users user = Users.find(" email = ? ", email).first();
			if(user==null){
				flash.put("error", "该用户不存在！请重新输入邮箱");
				login();
			}else{
				if(password==null||user.password==null||!user.password.equals(password)){
					flash.put("error", "邮箱跟密码不匹配，请重新输入！");
					login();
				}else{
					session.put("LOGIN_ID", user.id);
					session.put("LOGIN_NAME", user.name);
					UserOper.types();
				}
			}
		}
	}
	
	/**
	 * 语言转换
	 * @param c_N
	 */
	public static void language(Integer c_N){
		c_N = c_N==null?1:c_N;
		if(c_N==0){
			session.put("C_N", false);
		}else
			session.put("C_N", true);
		index();
	}
	

}