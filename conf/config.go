package conf

import (
	"gopkg.in/yaml.v3"
	"log"
	"os"
)

var (
	YAML = &struct {
		APP    *cApp
		System *cSystem
		Static *cStatic
		Baidu  *cBaidu
	}{}
)

type cApp struct {
	Id      string
	Debug   bool
	Version string
}

type cSystem struct {
	Url string
}

type cStatic struct {
	TemplatePath string
	OutputPath   string
}
type cBaidu struct {
	Url       string
	ApiKey    string
	SecretKey string
}

func Setup(filename string) {
	data, err := os.ReadFile(filename)
	if err != nil {
		log.Panicln(err)
	}

	err = yaml.Unmarshal(data, &YAML)
	if err != nil {
		log.Panicln(err)
	}
}

func App() *cApp {
	return YAML.APP
}

func System() *cSystem {
	return YAML.System
}

func Static() *cStatic {
	return YAML.Static
}
func Baidu() *cBaidu {
	return YAML.Baidu
}
