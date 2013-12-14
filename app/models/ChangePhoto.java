package models;

import javax.persistence.Entity;

import play.db.jpa.Model;

/**
 * 首页轮转图片
 * @author YangShanCheng
 */
@Entity
public class ChangePhoto extends Model {

	public String url; //图片路径
	
	public String name; // 产品名称
	
	public String eName; //英文名称
	
	public String content; //产品简介
	
	public String eContent; //英文介绍
	
}
