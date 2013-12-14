package models;

import javax.persistence.Entity;
import javax.persistence.Lob;
import javax.persistence.ManyToOne;

import play.db.jpa.Model;

/**
 * 商品信息
 * @author YangShanCheng
 */
@Entity
public class Product extends Model {

	public String name; //商品名称
	
	public String guige; // 商品规格
	
	@ManyToOne
	public ProductTypes productType; //所属类别
	
	public String zhong; //产品重量
	
	public String img; //商品图片
	
	public String openImg; //商品图片
	
	public String address; //生产地
	
	public String eName; //英文名称

}
