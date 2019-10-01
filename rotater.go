package main

import (
	"fmt"
	"regexp"
	"strings"
)

func FirstImage(s string) string {
	extStart := strings.LastIndex(s, ".")
	extension := s[extStart:]

	imagePathWithoutExt := strings.Split(s, "..")[0]
	imagePath := fmt.Sprintf("%s%s", imagePathWithoutExt, extension)

	return imagePath
}

func IntervalImage(s string) string {
	imgRangeRegexp := regexp.MustCompile("[\\d]+\\.\\.[\\d]+")
	imgRange := imgRangeRegexp.FindStringSubmatch(s)[0]

	indexLength := len(strings.Split(imgRange, "..")[0])
	imgPattern := imgRangeRegexp.ReplaceAllString(s, strings.Repeat("#", indexLength))
	imgPatternWithRange := fmt.Sprintf("%s|%s", imgPattern, imgRange)

	return imgPatternWithRange
}
