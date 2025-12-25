package xlsx

import (
	"errors"
	"reflect"
	"strconv"
	"time"
)

var (
	errNilInterface     = errors.New("nil pointer is not a valid argument")
	errNotStructPointer = errors.New("argument must be a pointer to struct")
	errInvalidTag       = errors.New(`invalid tag: must have the format xlsx:idx`)
)

// XLSXUnmarshaler is the interface implemented for types that can unmarshal a Row
// as a representation of themselves.
type XLSXUnmarshaler interface {
	Unmarshal(*Row) error
}

// ReadStruct 从 r 读取结构体到 ptr。接受一个指向结构体的指针。
// 此代码期望一个标签 xlsx:"N"，其中 N 是要使用的单元格索引。
// 支持基本类型如 int、string、float64 和 bool。
// 通过 parseValue 转换，也支持复杂类型如 map、slice、array。
func (r *Row) ReadStruct(ptr interface{}) error {
	if ptr == nil {
		return errNilInterface
	}
	//check if the type implements XLSXUnmarshaler. If so,
	//just let it do the work.
	unmarshaller, ok := ptr.(XLSXUnmarshaler)
	if ok {
		return unmarshaller.Unmarshal(r)
	}
	v := reflect.ValueOf(ptr)
	if v.Kind() != reflect.Ptr {
		return errNotStructPointer
	}
	v = v.Elem()
	if v.Kind() != reflect.Struct {
		return errNotStructPointer
	}
	n := v.NumField()
	for i := 0; i < n; i++ {
		field := v.Type().Field(i)
		idx := field.Tag.Get("xlsx")
		//do a recursive check for the field if it is a struct or a pointer
		//even if it doesn't have a tag
		//ignore if it has a - or empty tag
		isTime := false
		switch {
		case idx == "-":
			continue
		case field.Type.Kind() == reflect.Ptr || field.Type.Kind() == reflect.Struct:
			var structPtr interface{}
			if !v.Field(i).CanSet() {
				continue
			}
			if field.Type.Kind() == reflect.Struct {
				structPtr = v.Field(i).Addr().Interface()
			} else {
				structPtr = v.Field(i).Interface()
			}
			//check if the container is a time.Time
			_, isTime = structPtr.(*time.Time)
			if isTime {
				break
			}
			err := r.ReadStruct(structPtr)
			if err != nil {
				return err
			}
			continue
		case len(idx) == 0:
			continue
		}
		pos, err := strconv.Atoi(idx)
		if err != nil {
			return errInvalidTag
		}

		cell := r.GetCell(pos)
		if cell.Value == "" {
			continue
		}
		fieldV := v.Field(i)
		//continue if the field is not settable
		if !fieldV.CanSet() {
			continue
		}
		if isTime {
			t, err := cell.GetTime(false)
			if err != nil {
				return err
			}
			if field.Type.Kind() == reflect.Ptr {
				fieldV.Set(reflect.ValueOf(&t))
			} else {
				fieldV.Set(reflect.ValueOf(t))
			}
			continue
		}
		switch field.Type.Kind() {
		case reflect.String:
			value, err := cell.FormattedValue()
			if err != nil {
				return err
			}
			fieldV.SetString(value)
		case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
			value, err := cell.Int64()
			if err != nil {
				return err
			}
			fieldV.SetInt(value)
		case reflect.Float64:
			value, err := cell.Float()
			if err != nil {
				return err
			}
			fieldV.SetFloat(value)
		case reflect.Bool:
			value := cell.Bool()
			fieldV.SetBool(value)
		case reflect.Map, reflect.Slice, reflect.Array:
			err := r.setComplexType(cell, fieldV, field.Type)
			if err != nil {
				return err
			}
		}
	}
	return nil
}

// setComplexType 从单元格值设置复杂类型（map、slice、array），
// 使用 parseValue 进行转换。
func (r *Row) setComplexType(cell *Cell, fieldV reflect.Value, fieldType reflect.Type) error {
	value, err := cell.FormattedValue()
	if err != nil {
		return err
	}
	//使用 parseValue 将字符串转换为适当的类型
	parsedValue := parseValue(value)
	if parsedValue == nil {
		return nil
	}
	//将 parsedValue 转换为目标类型
	convertedValue, err := convertToReflectType(parsedValue, fieldType)
	if err != nil {
		return err
	}
	if convertedValue.IsValid() && convertedValue.Type().AssignableTo(fieldType) {
		fieldV.Set(convertedValue)
	}
	return nil
}

// convertToReflectType 将 interface{} 值转换为目标类型的 reflect.Value，
// 处理 map、slice 和 array 类型。
func convertToReflectType(value interface{}, targetType reflect.Type) (reflect.Value, error) {
	if value == nil {
		return reflect.Zero(targetType), nil
	}
	valueV := reflect.ValueOf(value)
	valueT := valueV.Type()
	switch targetType.Kind() {
	case reflect.Map:
		return convertToMap(valueV, targetType)
	case reflect.Slice:
		return convertToSlice(valueV, targetType)
	case reflect.Array:
		return convertToArray(valueV, targetType)
	default:
		//对于基本类型，尝试直接转换
		if valueT.AssignableTo(targetType) {
			return valueV, nil
		}
		if valueT.ConvertibleTo(targetType) {
			return valueV.Convert(targetType), nil
		}
		return reflect.Value{}, errors.New("cannot convert value to target type")
	}
}

// convertToMap 将值转换为目标类型的 map。
func convertToMap(valueV reflect.Value, targetType reflect.Type) (reflect.Value, error) {
	if valueV.Kind() != reflect.Map {
		return reflect.Value{}, errors.New("value is not a map")
	}
	//创建目标类型的新 map
	result := reflect.MakeMap(targetType)
	keyType := targetType.Key()
	elemType := targetType.Elem()
	//遍历源 map
	for _, key := range valueV.MapKeys() {
		elem := valueV.MapIndex(key)
		//转换 key
		convertedKey, err := convertValue(key, keyType)
		if err != nil {
			continue
		}
		//转换 value
		convertedElem, err := convertValue(elem, elemType)
		if err != nil {
			continue
		}
		result.SetMapIndex(convertedKey, convertedElem)
	}
	return result, nil
}

// convertToSlice 将值转换为目标类型的 slice。
func convertToSlice(valueV reflect.Value, targetType reflect.Type) (reflect.Value, error) {
	if valueV.Kind() != reflect.Slice && valueV.Kind() != reflect.Array {
		return reflect.Value{}, errors.New("value is not a slice or array")
	}
	elemType := targetType.Elem()
	length := valueV.Len()
	result := reflect.MakeSlice(targetType, 0, length)
	//遍历源 slice/array
	for i := 0; i < length; i++ {
		elem := valueV.Index(i)
		var convertedElem reflect.Value
		var err error
		switch elemType.Kind() {
		case reflect.Array, reflect.Slice:
			convertedElem, err = convertToReflectType(elem.Interface(), elemType)
		default:
			convertedElem, err = convertValue(elem, elemType)
		}
		if err != nil {
			continue
		}
		result = reflect.Append(result, convertedElem)
	}
	return result, nil
}

// convertToArray 将值转换为目标类型的 array。
func convertToArray(valueV reflect.Value, targetType reflect.Type) (reflect.Value, error) {
	if valueV.Kind() != reflect.Slice && valueV.Kind() != reflect.Array {
		return reflect.Value{}, errors.New("value is not a slice or array")
	}
	elemType := targetType.Elem()
	arrayLen := targetType.Len()
	result := reflect.New(targetType).Elem()
	sourceLen := valueV.Len()
	//复制元素，最多复制源长度和数组长度的最小值
	copyLen := sourceLen
	if copyLen > arrayLen {
		copyLen = arrayLen
	}
	for i := 0; i < copyLen; i++ {
		elem := valueV.Index(i)
		var convertedElem reflect.Value
		var err error
		switch elemType.Kind() {
		case reflect.Array, reflect.Slice:
			convertedElem, err = convertToReflectType(elem.Interface(), elemType)
		default:
			convertedElem, err = convertValue(elem, elemType)
		}
		if err != nil {
			continue
		}
		result.Index(i).Set(convertedElem)
	}
	return result, nil
}

// convertValue 将 reflect.Value 转换为目标类型。
func convertValue(valueV reflect.Value, targetType reflect.Type) (reflect.Value, error) {
	if !valueV.IsValid() {
		return reflect.Zero(targetType), nil
	}
	valueT := valueV.Type()
	//直接赋值
	if valueT.AssignableTo(targetType) {
		return valueV, nil
	}
	//类型转换
	if valueT.ConvertibleTo(targetType) {
		return valueV.Convert(targetType), nil
	}
	//通过提取底层值来处理 interface{} 类型
	if valueV.Kind() == reflect.Interface {
		return convertValue(valueV.Elem(), targetType)
	}
	return reflect.Value{}, errors.New("cannot convert value to target type")
}
