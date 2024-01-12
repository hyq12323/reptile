package util

import (
	"log/slog"
	"os"
	"strings"
)

// IsFileExist 判断目录下是否有同名Excel，有就删除
func IsFileExist(fileName string) (isExist bool, err error) {
	_, err = os.Stat(fileName)
	if err != nil {
		// 文件不存在
		if os.IsNotExist(err) {
			err = nil
			return
		}
		slog.Error("[IsFileExist]", "os.Stat", err)
		return
	}
	// 删除同名文件
	//err = os.Remove(fileName)
	//if err != nil {
	//	slog.Error("[IsFileExist]", "os.Remove", err)
	//	return
	//}
	isExist = true
	return
}

// ExtractTextInParentheses 获取括号中的文字
func ExtractTextInParentheses(input string) string {
	startIndex := strings.Index(input, "（")
	if startIndex == -1 {
		return "" // 未找到左括号
	}

	endIndex := strings.Index(input, "）")
	if endIndex == -1 || endIndex <= startIndex {
		return "" // 未找到右括号或右括号在左括号之前
	}

	return input[startIndex+3 : endIndex]
}
