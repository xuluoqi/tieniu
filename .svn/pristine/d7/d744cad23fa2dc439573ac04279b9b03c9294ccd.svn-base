package models;

import java.util.List;

import javax.persistence.Entity;
import javax.persistence.ManyToOne;
import javax.persistence.Transient;

import org.hibernate.annotations.NotFound;
import org.hibernate.annotations.NotFoundAction;

import play.db.jpa.Model;

/**
 * 产品类型
 * @author YangShanCheng
 */
@Entity
public class ProductTypes extends Model {

	public String name; //名称
	
	@ManyToOne
	@NotFound(action=NotFoundAction.IGNORE)
	public ProductTypes type; // 父类
	
	public String eName; //英文名称
	
	@Transient
	public List<ProductTypes> types; //每个类别下面的分类

	public ProductTypes(String name, ProductTypes type,String eName) {
		super();
		this.name = name;
		this.type = type;
		this.eName = eName;
	}

	public ProductTypes() {
		super();
		// TODO Auto-generated constructor stub
	}
	
	public ProductTypes(String name, Long id) {
		super();
		ProductTypes type = ProductTypes.findById(id);
		this.name = name;
		this.type = type;
	}
	
}
