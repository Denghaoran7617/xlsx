package xlsx

import (
	"encoding/json"
	"strconv"
	"strings"
)

// ------------------------ 数据类型解析 ----------------------------
// 自动判断类型：int64、float64、[]int64、[]string、map、slice...
// 优先级：JSON > int64 > float > 数组
func parseValue(s string) interface{} {
	// JSON 对象 / 数组
	if strings.HasPrefix(s, "{") || strings.HasPrefix(s, "[") {
		var v interface{}
		if json.Unmarshal([]byte(s), &v) == nil {
			return convertJSONNumber(v)
		}
	}

	// 尝试 int64
	if i64, err := strconv.ParseInt(s, 10, 64); err == nil {
		return i64
	}

	// 尝试 float
	if f, err := strconv.ParseFloat(s, 64); err == nil {
		return f
	}

	// 自定义解析数组：1,2,3
	if strings.Contains(s, ",") {
		return parseList(s)
	}

	return s
}

// 解析类似 "1,2,3" 或 "a,b,c"
func parseList(s string) interface{} {
	parts := strings.Split(s, ",")
	intArr := []int64{}
	strArr := []string{}

	isIntArray := true

	for _, p := range parts {
		p = strings.TrimSpace(p)
		if v, err := strconv.ParseInt(p, 10, 64); err == nil {
			intArr = append(intArr, v)
		} else {
			isIntArray = false
			strArr = append(strArr, p)
		}
	}

	if isIntArray {
		return intArr
	}
	return strArr
}

// JSON 数字会转为 float64，这里递归转换成更合适类型
func convertJSONNumber(v interface{}) interface{} {
	switch t := v.(type) {
	case map[string]interface{}:
		for k, val := range t {
			t[k] = convertJSONNumber(val)
		}
		return t

	case []interface{}:
		if len(t) == 0 {
			return t
		}
		// 先递归转换所有元素
		for i, val := range t {
			t[i] = convertJSONNumber(val)
		}
		// 尝试转换为具体类型的数组
		return convertSliceToTypedArray(t)

	case float64:
		// 判断是否是整数
		if float64(int64(t)) == t {
			return int64(t)
		}
		return t
	default:
		return v
	}
}

// 将 []interface{} 转换为具体的数组类型
func convertSliceToTypedArray(slice []interface{}) interface{} {
	if len(slice) == 0 {
		return slice
	}

	// 检查第一个元素的类型
	firstType := getElementType(slice[0])
	if firstType == typeOther {
		return slice // 无法确定类型，返回原数组
	}

	// 检查所有元素是否都是同一类型
	for i := 1; i < len(slice); i++ {
		if getElementType(slice[i]) != firstType {
			return slice // 类型不一致，返回原数组
		}
	}

	// 根据类型创建对应的数组
	switch firstType {
	case typeInt64:
		result := make([]int64, len(slice))
		for i, v := range slice {
			result[i] = v.(int64)
		}
		return result
	case typeFloat64:
		result := make([]float64, len(slice))
		for i, v := range slice {
			result[i] = v.(float64)
		}
		return result
	case typeString:
		result := make([]string, len(slice))
		for i, v := range slice {
			result[i] = v.(string)
		}
		return result
	case typeBool:
		result := make([]bool, len(slice))
		for i, v := range slice {
			result[i] = v.(bool)
		}
		return result
	default:
		return slice
	}
}

// 元素类型标识
type elementType int

const (
	typeInt64 elementType = iota
	typeFloat64
	typeString
	typeBool
	typeOther
)

// 获取元素的类型标识
func getElementType(v interface{}) elementType {
	switch v.(type) {
	case int64:
		return typeInt64
	case float64:
		return typeFloat64
	case string:
		return typeString
	case bool:
		return typeBool
	default:
		return typeOther
	}
}
