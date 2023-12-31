package main

import (
	"log/slog"
	"reptile/common/slogx"
	"reptile/conf"
	"reptile/service"
)

func init() {
	conf.Setup("reptile.yaml")
	slogx.SetDefault(conf.YAML.APP.Debug)
	slog.Info("注册会计信息系统爬虫程序", "版本号", conf.YAML.APP.Version)
}

func main() {
	service.Reptile()
}
