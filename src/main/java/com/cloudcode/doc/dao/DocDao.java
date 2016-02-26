package com.cloudcode.doc.dao;

import java.util.List;

import javax.annotation.Resource;

import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Transactional;

import com.cloudcode.doc.ProjectConfig;
import com.cloudcode.doc.model.Doc;
import com.cloudcode.framework.dao.BaseModelObjectDao;
import com.cloudcode.framework.dao.ModelObjectDao;
import com.cloudcode.framework.utils.HQLParamList;
import com.cloudcode.framework.utils.PageRange;
import com.cloudcode.framework.utils.PaginationSupport;

@Repository
public class DocDao extends BaseModelObjectDao<Doc> {
	
	@Resource(name = ProjectConfig.PREFIX + "docDao")
	private ModelObjectDao<Doc> docDao;
	
	@Transactional
	public void addDoc(Doc entity) {
		docDao.createObject(entity);
	}
	public PaginationSupport<Doc> queryPagingData(Doc hhXtCd, PageRange pageRange) {
		HQLParamList hqlParamList = new HQLParamList();
		List<Object> list=null;
		return this.queryPaginationSupport(Doc.class, hqlParamList, pageRange);
	}
}
