package util

import (
	"bytes"
	"encoding/json"
	"errors"
	"fmt"
	"io/ioutil"
	"log/slog"
	"net/http"
	"net/url"
	"strings"
)

var client = &http.Client{}

func HttpPostByJson(apiUrl string, params map[string]any, headers map[string]string) (respBody []byte, err error) {
	var body *bytes.Buffer
	var req *http.Request
	if params != nil {
		var b []byte
		b, err = json.Marshal(params)
		if err != nil {
			slog.Error("[HttpPost]", "json.Marshal", err)
			return
		}
		body = bytes.NewBuffer(b)
		req, err = http.NewRequest("POST", apiUrl, body)
	} else {
		req, err = http.NewRequest("POST", apiUrl, nil)
	}
	if err != nil {
		slog.Error("[HttpPost]", "http.NewRequest", err)
		return
	}
	if headers != nil {
		for k, v := range headers {
			req.Header.Set(k, v)
		}
	}
	req.Header.Set("Content-Type", "application/json")
	resp, err := client.Do(req)
	if err != nil {
		slog.Error("[HttpPost]", "client.Do", err)
		return
	}
	defer resp.Body.Close()
	respBody, err = ioutil.ReadAll(resp.Body)
	if resp.StatusCode != http.StatusOK {
		err = errors.New(fmt.Sprintf("response fail: %s", string(respBody)))
		slog.Error("[HttpPost]", "Code != 200", err)
		return
	}
	return
}

// HttpGet 发送 HTTP GET 请求并返回响应内容
func HttpGet(apiUrl string, params map[string]string, headers map[string]string) (respBody []byte, err error) {
	req, err := http.NewRequest("GET", apiUrl, nil)
	if err != nil {
		slog.Error("[HttpPost]", "http.NewRequest", err)
		return
	}
	if headers != nil {
		for k, v := range headers {
			req.Header.Set(k, v)
		}
	}
	if params != nil {
		q := req.URL.Query()
		for k, v := range params {
			q.Set(k, v)
		}
		req.URL.RawQuery = q.Encode()
	}
	resp, err := client.Do(req)
	if err != nil {
		slog.Error("[HttpPost]", "client.Do", err)
		return
	}
	defer resp.Body.Close()
	respBody, err = ioutil.ReadAll(resp.Body)
	if resp.StatusCode != http.StatusOK {
		err = errors.New(fmt.Sprintf("response fail: %s", string(respBody)))
		slog.Error("[HttpPost]", "Code != 200", err)
		return
	}
	return
}

func HttpPostByForm(apiUrl string, params map[string]string, headers map[string]string) (respBody []byte, err error) {
	// 将参数编码为字符串
	data := url.Values{}
	if params != nil {
		for key, value := range params {
			data.Set(key, value)
		}
	}
	payload := strings.NewReader(data.Encode())
	req, err := http.NewRequest("POST", apiUrl, payload)
	if err != nil {
		slog.Error("[HttpPost]", "http.NewRequest", err)
		return
	}
	if headers != nil {
		for k, v := range headers {
			req.Header.Set(k, v)
		}
	}
	req.Header.Set("Content-Type", "application/json")
	resp, err := client.Do(req)
	if err != nil {
		slog.Error("[HttpPost]", "client.Do", err)
		return
	}
	defer resp.Body.Close()
	respBody, err = ioutil.ReadAll(resp.Body)
	if resp.StatusCode != http.StatusOK {
		err = errors.New(fmt.Sprintf("response fail: %s", string(respBody)))
		slog.Error("[HttpPost]", "Code != 200", err)
		return
	}
	return
}
