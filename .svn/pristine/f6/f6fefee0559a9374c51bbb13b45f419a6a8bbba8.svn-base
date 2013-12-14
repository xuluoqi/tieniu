package models;

import javax.persistence.Entity;
import javax.persistence.Lob;

import play.db.jpa.Model;
import play.mvc.Scope.Session;

/**
 * 公司信息
 * 习惯了 用User吧
 * @author YangShanCheng
 */
@Entity
public class Users extends Model {

	public String name; //中文公司名称
	
	public String eName; //英文公司名称
	
	public String phone; //电话
	
	public String email; //邮箱
	
	public String address; //中文地址
	
	public String eAddress; //英文地址
	
	public String zips;  //邮编
	
	public String fax;  //传真
	
	public String contact; // 中文联系人
	
    public String eContact; // 英文联系人 	
	@Lob
	public String content; //公司简介
	
	@Lob
	public String eContent; //英文公司简介
	
	public String password; //密码
	
	
	/**
	 * 获取登陆用户
	 * 
	 * @param session
	 * @return
	 */
	public static Users getUser(Session session) {
		Users user = null;
		Long id = session.get("LOGIN_ID") == null ? 0 : Long.valueOf(session.get("LOGIN_ID").toString());
		if (id > 0) {
			user = Users.findById(id);
		}
		return user;
	}
}
