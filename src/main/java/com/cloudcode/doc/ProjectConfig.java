package com.cloudcode.doc;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.ComponentScan;
import org.springframework.context.annotation.PropertySource;

import com.cloudcode.doc.model.Doc;
import com.cloudcode.framework.annotation.ModuleConfig;
import com.cloudcode.framework.bean.ProjectBeanNameGenerator;
import com.cloudcode.framework.dao.ModelObjectDao;
import com.cloudcode.framework.dao.impl.BaseDaoImpl;



@ModuleConfig(name=ProjectConfig.NAME,domainPackages={"com.cloudcode.doc.model"})
@ComponentScan(basePackages={"com.cloudcode.doc.*"},nameGenerator=ProjectBeanNameGenerator.class)
@PropertySource(name="cloudcode.evn",value={"classpath:proj.properties"})
public class ProjectConfig {
	public static final String NAME="doc";
	public static final String PREFIX=NAME+".";

	@Bean(name=PREFIX+"docDao")
	public ModelObjectDao<Doc> generateDocDao(){
		return new BaseDaoImpl<Doc>(Doc.class);
	}
}
