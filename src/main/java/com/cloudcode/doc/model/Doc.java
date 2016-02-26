package com.cloudcode.doc.model;

import javax.persistence.Entity;
import javax.persistence.Table;

import com.cloudcode.doc.ProjectConfig;
import com.cloudcode.framework.model.BaseModelObject;

@Entity(name = ProjectConfig.NAME + "doc")
@Table(name = ProjectConfig.NAME + "_doc")
@org.hibernate.annotations.Entity(dynamicInsert = true, dynamicUpdate = true)
public class Doc extends BaseModelObject{

	/**
	 * 
	 */
	private static final long serialVersionUID = -4102270536834403612L;
	
	private String name;

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}
	
	
}
