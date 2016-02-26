package com.cloudcode.doc.mvc;

import javax.servlet.http.HttpServletRequest;
import javax.validation.Valid;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.validation.Validator;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.ModelAndView;

import com.cloudcode.doc.dao.DocDao;
import com.cloudcode.doc.model.Doc;
import com.cloudcode.doc.utils.WordToHtml;
import com.cloudcode.framework.controller.CrudController;

@Controller
@RequestMapping("/doc")
public class DocController extends CrudController<Doc> {
	@Autowired
	private DocDao docDao;
	
	@RequestMapping(value = "futuresTypeList")
	public ModelAndView futuresTypeList() {
		ModelAndView modelAndView = new ModelAndView();
		modelAndView.setViewName("classpath:com/cloudcode/doc/ftl/list.ftl");
		modelAndView.addObject("result", "cloudcode");
		return modelAndView;
	}

	@RequestMapping(value = "create")
	public ModelAndView create() {
		ModelAndView modelAndView = new ModelAndView();
		modelAndView.setViewName("classpath:com/cloudcode/doc/ftl/detail.ftl");
		modelAndView.addObject("result", "cloudcode");
		modelAndView.addObject("entityAction", "create");
		return modelAndView;
	}

	@RequestMapping(value = "/createFuturesType", method = { RequestMethod.POST,
			RequestMethod.GET })
	public @ResponseBody void createFuturesType(@ModelAttribute  @Valid Doc doc, HttpServletRequest request) {
		String text = "C://test//test.doc";
		String html="C://test//test.html";
		WordToHtml.wordToHtml(text, html);
	}
	
	@Override
	protected Validator getValidator() {
		// TODO Auto-generated method stub
		return null;
	}
	
}
