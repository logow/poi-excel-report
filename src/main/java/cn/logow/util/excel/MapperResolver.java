package cn.logow.util.excel;

import java.beans.BeanInfo;
import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

class MapperResolver {
	
	private static final String DEFAULT_MAPPER_FILE = "/ExcelMapper.xml";
	
	private static Map<String, ColumnSpec[]> mapperCache = new HashMap<String, ColumnSpec[]>();
	
	private static final Log log = LogFactory.getLog(MapperResolver.class);
	
	static {
		InputStream xmlStream = MapperResolver.class.getResourceAsStream(DEFAULT_MAPPER_FILE);
		if (xmlStream != null) {
			log.info("Loading excel mappers from " + DEFAULT_MAPPER_FILE);
			try {
				readXMLMappers(xmlStream);
			} catch (Exception e) {
				log.error("Error loading excel mappers", e);
			} finally {
				try {
					xmlStream.close();
				} catch (IOException e) {
					log.error("Error closing resource", e);
				}
			}
		}
	}
	
	public static ColumnSpec[] resolveColumns(Class<?> entityClass) {
		ColumnSpec[] specs = mapperCache.get(entityClass.getName());
		if (specs == null) {
			synchronized (MapperResolver.class) {
				if (!mapperCache.containsKey(entityClass.getName())) {
					specs = ColumnResolver.getInstance().resolveColumns(entityClass);
					mapperCache.put(entityClass.getName(), specs);
				}
			}
		}
		return specs;
	}

	private static void readXMLMappers(InputStream xmlStream) throws Exception {
		SAXReader xmlReader = new SAXReader();
		Document doc = xmlReader.read(xmlStream);
		Element root = doc.getRootElement();
		List<?> mappers = root.elements("mapper");
		for (Object mapper : mappers) {
			loadXMLMapper((Element) mapper);
		}
	}
	
	private static void loadXMLMapper(Element mapper) throws Exception {
		String className = mapper.attributeValue("class");
		if (className == null) {
			return;
		}
		
		Class<?> mapperClass = Class.forName(className);
		Map<String, Method> getters = listGetters(mapperClass);
		List<ColumnSpec> list = new ArrayList<ColumnSpec>();
		List<?> eles = mapper.elements("column");
		for (Object ele : eles) {
			ColumnSpec spec = parseXMLColumn((Element) ele);
			if (spec != null && getters.containsKey(spec.name)) {
				spec.getter = getters.get(spec.name);
				spec.setDefaultFormat();
				list.add(spec);
				if (log.isDebugEnabled()) {
					log.debug(spec);
				}
			}
		}
		
		if (!list.isEmpty()) {
			mapperCache.put(className, list.toArray(new ColumnSpec[list.size()]));
			log.info("Excel mapper resolved: " + className);
		}
	}
	
	private static ColumnSpec parseXMLColumn(Element ele) {
		String name = ele.attributeValue("name");
		String title = ele.attributeValue("title");
		String width = ele.attributeValue("width");
		String align = ele.attributeValue("align");
		String format = ele.attributeValue("format");
		return ColumnSpec.createSpec(name, title, width, align, format);
	}
	
	private static Map<String, Method> listGetters(Class<?> beanClass) {
		try {
			Map<String, Method> getters = new HashMap<String, Method>();
			BeanInfo beanInfo = Introspector.getBeanInfo(beanClass);
			PropertyDescriptor[] pds = beanInfo.getPropertyDescriptors();
			if (pds != null) {
				for (PropertyDescriptor pd : pds) {
					if (pd.getReadMethod() != null) {
						getters.put(pd.getName(), pd.getReadMethod());
					}
				}
			}
			return getters;
		} catch (IntrospectionException e) {
			throw new RuntimeException(e);
		}
	}
}
