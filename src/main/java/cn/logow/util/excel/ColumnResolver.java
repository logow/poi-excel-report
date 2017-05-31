package cn.logow.util.excel;

import java.beans.BeanInfo;
import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;

public class ColumnResolver {
	
	private static ColumnResolver instance = new ColumnResolver();
	
	public static ColumnResolver getInstance() {
		return instance;
	}

	public ColumnSpec[] resolveColumns(Class<?> entityClass) {
		List<Field> fields = getColumnFields(entityClass);
		if (!fields.isEmpty()) {
			return resolveColumnFields(fields);
		} else {
			return resolveColumnProperties(getColumnProperties(entityClass));
		}
	}
	
	private void sortColumns(List<ColumnSpec> columns) {
		Collections.sort(columns, new Comparator<ColumnSpec>(){
			@Override
			public int compare(ColumnSpec o1, ColumnSpec o2) {
				if (o1.order < o2.order) {
					return -1;
				} else if (o1.order > o2.order) {
					return 1;
				} else {
					return 0;
				}
			}
		});
	}
	
	private ColumnSpec[] resolveColumnProperties(List<PropertyDescriptor> props) {
		List<ColumnSpec> specs = new ArrayList<ColumnSpec>(props.size());
		for (PropertyDescriptor pd : props) {
			Column col = pd.getReadMethod().getAnnotation(Column.class);
			if (col != null) {
				ColumnSpec spec = ColumnSpec.createSpec(pd.getName(), col);
				spec.getter = pd.getReadMethod();
				spec.setDefaultFormat();
				specs.add(spec);
			}
		}
		sortColumns(specs);
		return specs.toArray(new ColumnSpec[specs.size()]);
	}

	private ColumnSpec[] resolveColumnFields(List<Field> fields) {
		List<ColumnSpec> specs = new ArrayList<ColumnSpec>(fields.size());
		for (Field f : fields) {
			Column col = f.getAnnotation(Column.class);
			if (col != null) {
				ColumnSpec spec = ColumnSpec.createSpec(f.getName(), col);
				spec.field = f;
				spec.setDefaultFormat();
				specs.add(spec);
			}
		}
		sortColumns(specs);
		return specs.toArray(new ColumnSpec[specs.size()]);
	}

	private List<Field> getColumnFields(Class<?> clazz) {
		Field[] fields = clazz.getDeclaredFields();
		List<Field> list = new ArrayList<Field>();
		for (Field field : fields) {
			if (isColumnField(field)) {
				forceAccessable(field);
				list.add(field);
			}
		}
		return list;
	}
	
	private List<PropertyDescriptor> getColumnProperties(Class<?> clazz) {
		List<PropertyDescriptor> props = new ArrayList<PropertyDescriptor>();
		try {
			BeanInfo beanInfo = Introspector.getBeanInfo(clazz);
			PropertyDescriptor[] pds = beanInfo.getPropertyDescriptors();
			if (pds != null) {
				for (PropertyDescriptor pd : pds) {
					Method getter = pd.getReadMethod();
					if (getter != null && getter.isAnnotationPresent(Column.class)) {
						props.add(pd);
					}
				}
			}
		} catch (IntrospectionException e) {
			throw new RuntimeException(e);
		}
		return props;
	}
	
	private boolean isColumnField(Field field) {
		int modifiers = field.getModifiers();
		if (Modifier.isStatic(modifiers) ||
			Modifier.isFinal(modifiers) ||
			Modifier.isTransient(modifiers)) {
			return false;
		}
		if (field.isAnnotationPresent(Column.class)) {
			return true;
		}
		return false;
	}
	
	private void forceAccessable(Field field) {
		if (!field.isAccessible()) {
			field.setAccessible(true);
		}
	}
}
