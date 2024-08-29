package utils

import (
	"strconv"
)

type Integer interface {
	int | int8 | int16 | int32 | int64 | uint | uint8 | uint16 | uint32 | uint64
}

func StrToInt[T Integer](str string) T {
	s, _ := strconv.Atoi(str)
	return T(s)
}
