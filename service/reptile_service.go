package service

import (
	"encoding/json"
	"fmt"
	"github.com/tealeg/xlsx"
	"github.com/xuri/excelize/v2"
	"log/slog"
	"reflect"
	"reptile/common/util"
	"reptile/conf"
	"reptile/model/reptile_model"
	"reptile/model/reptile_model/reply"
	"strconv"
	"strings"
	"sync"
	"time"
)

// RegisterInfos 事务所数据
var RegisterInfos []reptile_model.OfficeDetail

// PartnerInfos 合伙人数据
var PartnerInfos []reptile_model.PartnerInfo

// AccountantInfos 注册会计师数据
var AccountantInfos []reptile_model.AccountantInfo

// PractitionerInfos 从业人员数据
var PractitionerInfos []reptile_model.PractitionerInfo

// RegisterBranchInfos 分所事务所数据
var RegisterBranchInfos []reptile_model.FirmBranchInfo

// AccountantBranchInfos 分所注册会计师数据
var AccountantBranchInfos []reptile_model.AccountantInfo

// PractitionerBranchInfos 分所从业人员数据
var PractitionerBranchInfos []reptile_model.PractitionerInfo

// FirmDetailInfos 事务所详细数据
var FirmDetailInfos []reptile_model.FirmDetail

// Token 百度API的Token
var Token string

// RetryCount 重试次数
var RetryCount = []int{1, 2, 3}

// Reptile 主程序
func Reptile() {
	// 获取百度Token
	token, err := getBaiduToken()
	if err != nil {
		slog.Error("[Reptile]", "getBaiduToken", err)
		return
	}
	Token = token
	infos, err := readExcel()
	if err != nil {
		slog.Error("[Reptile]", "readExcel", err)
		return
	}
	for i, info := range infos {
		slog.Info("[Reptile]", "当前处理中的事务所", info.Name, "当前处理中序号", i+1, "总数", len(infos))
		func(req reply.ReadExcelReply) {
			var registerInfos reptile_model.OfficeDetail
			var firmDetail reptile_model.FirmDetail
			var partnerInfos []reptile_model.PartnerInfo
			var accountantInfos []reptile_model.AccountantInfo
			var practitionerInfos []reptile_model.PractitionerInfo
			var branchInfos reply.BranchReply
			// 事务所ID，用于后续接口调用
			var offGuid string
			// 事务所分所数据,用于后续接口调用
			var branchInfo []reply.FirmBranchInfo
			// 获取事务所数据
			registerInfos, firmDetail, offGuid, branchInfo, err = getFirmInfos(req)
			if err != nil {
				return
			}
			time.Sleep(400 * time.Millisecond)
			// 获取事务所合伙人数据
			partnerInfos, err = getPartnerInfos(offGuid, registerInfos.PartnerCount, req.Name)
			if err != nil {
				return
			}
			time.Sleep(400 * time.Millisecond)
			// 获取注册会计师数据
			accountantInfos, err = getAccountantInfos(offGuid, registerInfos.OffName, registerInfos.CpaNum)
			if err != nil {
				return
			}
			time.Sleep(400 * time.Millisecond)
			// 获取从业人员数据
			practitionerInfos, err = getPractitionerInfos(offGuid, registerInfos.OndutySum, req.Name)
			if err != nil {
				return
			}
			time.Sleep(400 * time.Millisecond)
			// 获取每个事务所分所的所有数据
			branchInfos, err = getFirmBranchInfos(branchInfo)
			if err != nil {
				return
			}
			RegisterInfos = append(RegisterInfos, registerInfos)
			PartnerInfos = append(PartnerInfos, partnerInfos...)
			AccountantInfos = append(AccountantInfos, accountantInfos...)
			PractitionerInfos = append(PractitionerInfos, practitionerInfos...)
			RegisterBranchInfos = append(RegisterBranchInfos, branchInfos.FirmBranchInfos...)
			AccountantBranchInfos = append(AccountantBranchInfos, branchInfos.BranchAccountantInfos...)
			PractitionerBranchInfos = append(PractitionerBranchInfos, branchInfos.BranchPractitionerInfo...)
			FirmDetailInfos = append(FirmDetailInfos, firmDetail)
			// 等待 0.5 秒
			time.Sleep(500 * time.Millisecond)
		}(info)
	}
	err = outPutExcel()
	if err != nil {
		slog.Error("[Reptile]", "outPutExcel", err)
		return
	}
}

// 获取事务所信息
func getFirmInfos(req reply.ReadExcelReply) (registerInfos reptile_model.OfficeDetail, firmDetail reptile_model.FirmDetail,
	offGuid string, branchInfo []reply.FirmBranchInfo, err error) {
	codeResp := reply.CodeReply{}
	codeResp, err = getPlatformCode("")
	if err != nil {
		slog.Error("[getFirmInfos]", "getPlatformCode", err)
		return
	}
	var text string
	for i := range RetryCount {
		text = getImageText(codeResp)
		if text != "" {
			break
		}
		if i == 2 {
			slog.Error("getFirmInfos", req.Name, "基础数据获取失败")
			return
		}
		// 验证码解析失败，再次获取验证码
		codeResp, err = getPlatformCode("")
		if err != nil {
			slog.Error("[getFirmInfos]", "getPlatformCode", err)
			return
		}
		// 等待 0.5 秒
		time.Sleep(500 * time.Millisecond)
	}
	officeParam := map[string]any{
		"offName":     req.Name,
		"ascGuid":     getAscGuid(req.Address),
		"offAllcode":  req.No,
		"verifyId":    codeResp.VerifyId,
		"verifyCode":  text,
		"currentPage": 1,
		"pageSize":    10,
	}
	officeUri := "publicQuery/getOfficeList"
	officeUrl := conf.System().Url + officeUri
	officeReply := reply.OfficeInfo{}
	var respBody []byte
	respBody, err = util.HttpPostByJson(officeUrl, officeParam, nil)
	if err != nil {
		return
	}
	err = json.Unmarshal(respBody, &officeReply)
	if err != nil {
		slog.Info("[getFirmInfos]", req.Name, "事务所列表数据获取失败")
		return
	}
	if len(officeReply.Result.List) == 0 {
		slog.Info("[getFirmInfos]", req.Name, "事务所列表为空")
		return
	}
	officeDetailParam := map[string]any{
		"offCode": officeReply.Result.List[0].OffAllCode,
	}
	officeDetailUri := "publicQuery/getOfficeDetailInfo"
	officeDetailUrl := conf.System().Url + officeDetailUri
	var detailReply reply.RegisterInfoReply
	var detailBody []byte
	for i := range RetryCount {
		detailBody, err = util.HttpPostByJson(officeDetailUrl, officeDetailParam, nil)
		if err != nil {
			if i == 2 {
				slog.Error("getFirmInfos", "util.HttpPostByJson", err)
				return
			}
			time.Sleep(400 * time.Millisecond)
			continue
		}
		err = json.Unmarshal(detailBody, &detailReply)
		if err != nil {
			if i == 2 {
				slog.Error("getFirmInfos", req.Name, "基础数据获取失败")
				return
			}
			time.Sleep(400 * time.Millisecond)
			continue
		}
		break
	}
	registerInfos = reptile_model.OfficeDetail{
		OffName:       detailReply.Info.HeadInfo.OffName,
		SubCount:      detailReply.Info.HeadInfo.SubCount,
		PartnerCount:  detailReply.Info.HeadInfo.PartnerCount,
		CpaNum:        detailReply.Info.HeadInfo.CpaNum,
		OndutySum:     detailReply.Info.HeadInfo.OndutySum,
		AllPerCount:   detailReply.Info.HeadInfo.AllPerCount,
		NoAllPerCount: detailReply.Info.HeadInfo.NoAllPerCount,
	}
	offType := util.ExtractTextInParentheses(detailReply.Info.HeadInfo.OffName)
	firmDetail = reptile_model.FirmDetail{
		OffName:     detailReply.Info.HeadInfo.OffName,
		OffCode:     detailReply.Info.HeadInfo.OffCode,
		OffType:     offType,
		RegMoney:    detailReply.Info.HeadInfo.RegMoney,
		AccountName: detailReply.Info.HeadInfo.AccountName,
		PassWord:    detailReply.Info.HeadInfo.PassWord,
		PassTime:    detailReply.Info.HeadInfo.PassTime,
		SubCount:    detailReply.Info.HeadInfo.SubCount,
		CpaNum:      detailReply.Info.HeadInfo.CpaNum,
		PhoneDecode: detailReply.Info.HeadInfo.PhoneDecode,
		Fax:         detailReply.Info.HeadInfo.Fax,
		OfficeAddr:  detailReply.Info.HeadInfo.OfficeAddr,
	}
	offGuid = officeReply.Result.List[0].OffAllCode
	branchInfo = detailReply.Info.SubOfficeList
	return
}

// 获取合伙人信息
func getPartnerInfos(offGuid, partnerCountStr, name string) (partnerInfos []reptile_model.PartnerInfo, err error) {
	// 获取股东总数
	var partnerCount int
	partnerCount, err = strconv.Atoi(partnerCountStr)
	if err != nil {
		partnerCount = 10
	}
	uri := fmt.Sprintf("publicQuery/getPartnerListByPage?offAllcode=%s&pageNow=%d&pageSize=%d", offGuid, 1, partnerCount)
	url := conf.System().Url + uri
	resp := reply.PartnerInfoReply{}
	var respBody []byte
	for i := range RetryCount {
		respBody, err = util.HttpPostByJson(url, nil, nil)
		if err != nil {
			if i == 2 {
				slog.Error("getPartnerInfos", "util.HttpPostByJson", err)
				return
			}
			time.Sleep(400 * time.Millisecond)
			continue
		}
		err = json.Unmarshal(respBody, &resp)
		if err != nil {
			if i == 2 {
				slog.Error("getPartnerInfos", name, "合伙人数据获取失败")
				return
			}
			time.Sleep(400 * time.Millisecond)
			continue
		}
		break
	}
	for i, row := range resp.Info.ResultMap.Rows {
		isCpa := "是"
		if row.IsCPA == "0" {
			isCpa = "否"
		}
		partnerInfos = append(partnerInfos, reptile_model.PartnerInfo{
			OffName: resp.Info.OffName,
			Number:  strconv.Itoa(i + 1),
			PerName: row.PerName,
			IsCPA:   isCpa,
			PerCode: row.PerCode,
		})
	}
	return
}

// 获取注册会计师信息
func getAccountantInfos(offGuid, offName, CpaNumStr string) (accountantInfos []reptile_model.AccountantInfo, err error) {
	var cpaNum int
	cpaNum, err = strconv.Atoi(CpaNumStr)
	if err != nil {
		cpaNum = 10
	}
	param := map[string]any{
		"offCode":     offGuid,
		"currentPage": 1,
		"pageSize":    cpaNum,
		"strAge":      "",
		"stuexpCode":  "",
	}
	uri := "publicQuery/getCpaList"
	url := conf.System().Url + uri
	resp := reply.AccountantInfoReply{}
	var respBody []byte
	for i := range RetryCount {
		respBody, err = util.HttpPostByJson(url, param, nil)
		if err != nil {
			if i == 2 {
				slog.Error("getAccountantInfos", "util.HttpPostByJson", err)
				return
			}
			time.Sleep(400 * time.Millisecond)
			continue
		}

		err = json.Unmarshal(respBody, &resp)
		if err != nil {
			if i == 2 {
				slog.Error("getAccountantInfos", offName, "注册会计师数据获取失败", err)
				return
			}
			time.Sleep(400 * time.Millisecond)
			continue
		}
		break
	}
	for i, info := range resp.Info.List {
		accountantInfos = append(accountantInfos, reptile_model.AccountantInfo{
			OffName: offName,
			Number:  strconv.Itoa(i + 1),
			PerName: info.PerName,
			PerCode: info.PerCode,
			Gender:  info.Gender,
			RegWord: info.RegWord,
		})
	}
	return
}

// 获取从业人员信息
func getPractitionerInfos(offGuid, ondutySumStr, name string) (practitionerInfos []reptile_model.PractitionerInfo, err error) {
	var ondutySum int
	ondutySum, err = strconv.Atoi(ondutySumStr)
	if err != nil {
		ondutySum = 10
	}

	uri := fmt.Sprintf("publicQuery/getEmployeeListByPage?offAllcode=%s&pageNow=%d&pageSize=%d", offGuid, 1, ondutySum)
	url := conf.System().Url + uri
	resp := reply.PractitionerInfoReply{}
	var respBody []byte
	for i := range RetryCount {
		respBody, err = util.HttpPostByJson(url, nil, nil)
		if err != nil {
			if err != nil {
				if i == 2 {
					slog.Error("getPractitionerInfos", "util.HttpPostByJson", err)
					return
				}
				time.Sleep(400 * time.Millisecond)
				continue
			}
		}

		err = json.Unmarshal(respBody, &resp)
		if err != nil {
			if i == 2 {
				slog.Error("getPractitionerInfos", name, "从业人员数据获取失败")
				return
			}
			time.Sleep(400 * time.Millisecond)
			continue
		}
		break
	}
	for i, info := range resp.Info.ResultMap.Rows {
		gender := "男"
		isPact, isSafety, isCpm := "是", "是", "是"
		if info.Gender == "2" {
			gender = "女"
		}
		if info.IsPact == "0" {
			isPact = "否"
		}
		if info.IsSafety == "0" {
			isSafety = "否"
		}
		if info.IsCpm == "0" {
			isCpm = "否"
		}
		practitionerInfos = append(practitionerInfos, reptile_model.PractitionerInfo{
			OffName:  resp.Info.OffName,
			Number:   strconv.Itoa(i + 1),
			EmpName:  info.EmpName,
			Gender:   gender,
			IntoTime: info.IntoTime,
			IsPact:   isPact,
			IsSafety: isSafety,
			IsCpm:    isCpm,
		})
	}
	return
}

// 获取事务所分所数据
func getFirmBranchInfos(branchInfo []reply.FirmBranchInfo) (firmBranchInfos reply.BranchReply, err error) {
	// 单个事务所的分所岗位人数数量数据
	var firmBranchInfo []reptile_model.FirmBranchInfo
	var accountantBranchInfo []reptile_model.AccountantInfo
	var practitionerBranchInfo []reptile_model.PractitionerInfo
	// 创建互斥锁
	var mu sync.Mutex
	var wg sync.WaitGroup
	maxConcurrent := 1 // 设置最大并发数
	infoChan := make(chan reply.FirmBranchInfo, maxConcurrent)
	for _, info := range branchInfo {
		infoChan <- info
		wg.Add(1)
		go func(req reply.FirmBranchInfo) {
			defer wg.Done()
			// 获取各个分所数量信息
			firmParam := map[string]any{
				"offCode": req.OffCode,
			}
			firmUrl := conf.System().Url + "publicQuery/getOfficeDetailInfo"
			var firmBody []byte
			firmResp := reply.FirmBranchInfoReply{}
			for i := range RetryCount {
				firmBody, err = util.HttpPostByJson(firmUrl, firmParam, nil)
				if err != nil {
					if i == 2 {
						slog.Error("getFirmBranchInfos", "util.HttpPostByJson", err)
						return
					}
					time.Sleep(400 * time.Millisecond)
					continue
				}
				err = json.Unmarshal(firmBody, &firmResp)
				if err != nil {
					if i == 2 {
						slog.Error("getFirmBranchInfos", req.OffName, "基础数据获取失败")
						return
					}
					time.Sleep(400 * time.Millisecond)
					continue
				}
				break
			}
			time.Sleep(400 * time.Millisecond)
			// 获取各个分所注册会计师信息
			var cpaNum int
			cpaNum, err = strconv.Atoi(firmResp.Info.HeadInfo.CpaNum)
			if err != nil {
				cpaNum = 10
			}
			accountantParam := map[string]any{
				"offCode":     req.OffCode,
				"currentPage": 1,
				"pageSize":    cpaNum,
				"strAge":      "",
				"stuexpCode":  "",
			}
			accountantUrl := conf.System().Url + "publicQuery/getCpaList"
			var accountantBody []byte
			accountantResp := reply.AccountantInfoReply{}
			for i := range RetryCount {
				accountantBody, err = util.HttpPostByJson(accountantUrl, accountantParam, nil)
				if err != nil {
					if i == 2 {
						slog.Error("getFirmBranchInfos", "util.HttpPostByJson", err)
						return
					}
					time.Sleep(400 * time.Millisecond)
					continue
				}
				err = json.Unmarshal(accountantBody, &accountantResp)
				if err != nil {
					if i == 2 {
						slog.Error("getFirmBranchInfos", req.OffName, "注册会计师数据获取失败")
						return
					}
					time.Sleep(400 * time.Millisecond)
					continue
				}
				break
			}
			time.Sleep(400 * time.Millisecond)
			var accountantInfos []reptile_model.AccountantInfo
			for i, list := range accountantResp.Info.List {
				accountantInfos = append(accountantInfos, reptile_model.AccountantInfo{
					OffName: req.OffName,
					Number:  strconv.Itoa(i + 1),
					PerName: list.PerName,
					PerCode: list.PerCode,
					Gender:  list.Gender,
					RegWord: list.RegWord,
				})
			}
			// 获取各个分所从业人员信息
			var ondutySum int
			ondutySum, err = strconv.Atoi(firmResp.Info.HeadInfo.OndutySum)
			if err != nil {
				ondutySum = 10
			}
			practitionerUri := fmt.Sprintf("publicQuery/getEmployeeListByPage?offAllcode=%s&pageNow=%d&pageSize=%d", req.OffCode, 1, ondutySum)
			practitionerUrl := conf.System().Url + practitionerUri
			var practitionerBody []byte
			practitionerResp := reply.PractitionerInfoReply{}
			for i := range RetryCount {
				practitionerBody, err = util.HttpPostByJson(practitionerUrl, nil, nil)
				if err != nil {
					if i == 2 {
						slog.Error("getFirmBranchInfos", "util.HttpPostByJson", err)
						return
					}
					time.Sleep(400 * time.Millisecond)
					continue
				}
				err = json.Unmarshal(practitionerBody, &practitionerResp)
				if err != nil {
					if i == 2 {
						slog.Error("getFirmBranchInfos", req.OffName, "从业人员数据获取失败")
						return
					}
					time.Sleep(400 * time.Millisecond)
					continue
				}
				break
			}
			var practitionerInfos []reptile_model.PractitionerInfo
			for i, row := range practitionerResp.Info.ResultMap.Rows {
				gender := "男"
				isPact, isSafety, isCpm := "是", "是", "是"
				if row.Gender == "2" {
					gender = "女"
				}
				if row.IsPact == "0" {
					isPact = "否"
				}
				if row.IsSafety == "0" {
					isSafety = "否"
				}
				if row.IsCpm == "0" {
					isCpm = "否"
				}
				practitionerInfos = append(practitionerInfos, reptile_model.PractitionerInfo{
					OffName:  req.OffName,
					Number:   strconv.Itoa(i + 1),
					EmpName:  row.EmpName,
					Gender:   gender,
					IntoTime: row.IntoTime,
					IsPact:   isPact,
					IsSafety: isSafety,
					IsCpm:    isCpm,
				})
			}
			// 加锁集中处理数据
			mu.Lock()
			firmBranchInfo = append(firmBranchInfo, firmResp.Info.HeadInfo)
			accountantBranchInfo = append(accountantBranchInfo, accountantInfos...)
			practitionerBranchInfo = append(practitionerBranchInfo, practitionerInfos...)
			mu.Unlock()
			// 通道移除
			<-infoChan
		}(info)
	}
	wg.Wait()
	firmBranchInfos.FirmBranchInfos = firmBranchInfo
	firmBranchInfos.BranchAccountantInfos = accountantBranchInfo
	firmBranchInfos.BranchPractitionerInfo = practitionerBranchInfo
	return
}

// 输出Excel
func outPutExcel() (err error) {
	// 生成事务所excel
	firmPath := conf.Static().OutputPath + "/" + "事务所.xlsx"
	isExist, err := util.IsFileExist(firmPath)
	if err != nil {
		slog.Error("[outPutExcel]", "util.IsFileExist", err)
		return
	}
	fireTitle := []string{"会计师事务所名称", "分所数量", "合伙人或股东人数", "注册会计师人数", "从业人员数量", "注册会计师人数（含分所）", "从业人员人数（含分所）"}
	err = writeExcel(isExist, firmPath, fireTitle, RegisterInfos)
	if err != nil {
		slog.Error("[outPutExcel]", "writeExcel", err)
		return
	}
	// 生成事务所合伙人信息
	partnerPath := conf.Static().OutputPath + "/" + "合伙人.xlsx"
	isExist, err = util.IsFileExist(partnerPath)
	if err != nil {
		slog.Error("[outPutExcel]", "util.IsFileExist", err)
		return
	}
	partnerTitle := []string{"会计师事务所名称", "序号", "合伙人（股东）姓名", "是否注师", "注师编号"}
	err = writeExcel(isExist, partnerPath, partnerTitle, PartnerInfos)
	if err != nil {
		slog.Error("[outPutExcel]", "writeExcel", err)
		return
	}
	// 生成注册会计师信息
	accountantPath := conf.Static().OutputPath + "/" + "注册会计师.xlsx"
	isExist, err = util.IsFileExist(accountantPath)
	if err != nil {
		slog.Error("[outPutExcel]", "util.IsFileExist", err)
		return
	}
	accountantTitle := []string{"会计师事务所名称", "序号", "姓名", "人员编号", "性别", "考核批准文号"}
	err = writeExcel(isExist, accountantPath, accountantTitle, AccountantInfos)
	if err != nil {
		slog.Error("[outPutExcel]", "writeExcel", err)
		return
	}

	// 生成从业人员信息
	practitionerPath := conf.Static().OutputPath + "/" + "从业人员.xlsx"
	isExist, err = util.IsFileExist(practitionerPath)
	if err != nil {
		slog.Error("[outPutExcel]", "util.IsFileExist", err)
		return
	}
	practitionerTitle := []string{"会计师事务所名称", "序号", "姓名", "性别", "进所时间", "是否签合同", "是否参加社保", "是否党员"}
	err = writeExcel(isExist, practitionerPath, practitionerTitle, PractitionerInfos)
	if err != nil {
		slog.Error("[outPutExcel]", "writeExcel", err)
		return
	}
	// 生成事务所分所excel
	firmBranchPath := conf.Static().OutputPath + "/" + "事务所（分所）.xlsx"
	isExist, err = util.IsFileExist(firmBranchPath)
	if err != nil {
		slog.Error("[outPutExcel]", "util.IsFileExist", err)
		return
	}
	fireBranchTitle := []string{"会计师事务所名称", "注册会计师总数", "从业人员总数"}
	err = writeExcel(isExist, firmBranchPath, fireBranchTitle, RegisterBranchInfos)
	if err != nil {
		slog.Error("[outPutExcel]", "writeExcel", err)
		return
	}
	// 生成分所注册会计师信息
	accountantBranchPath := conf.Static().OutputPath + "/" + "注册会计师（分所）.xlsx"
	isExist, err = util.IsFileExist(accountantBranchPath)
	if err != nil {
		slog.Error("[outPutExcel]", "util.IsFileExist", err)
		return
	}
	err = writeExcel(isExist, accountantBranchPath, accountantTitle, AccountantBranchInfos)
	if err != nil {
		slog.Error("[outPutExcel]", "writeExcel", err)
		return
	}
	// 生成分所从业人员信息
	practitionerBranchPath := conf.Static().OutputPath + "/" + "从业人员（分所）.xlsx"
	isExist, err = util.IsFileExist(practitionerBranchPath)
	if err != nil {
		slog.Error("[outPutExcel]", "util.IsFileExist", err)
		return
	}
	err = writeExcel(isExist, practitionerBranchPath, practitionerTitle, PractitionerBranchInfos)
	if err != nil {
		slog.Error("[outPutExcel]", "writeExcel", err)
		return
	}
	// 生成详细事务所信息
	firmDetailPath := conf.Static().OutputPath + "/" + "事务所基本信息.xlsx"
	isExist, err = util.IsFileExist(firmDetailPath)
	if err != nil {
		slog.Error("[outPutExcel]", "util.IsFileExist", err)
		return
	}
	fireDetailTitle := []string{"会计师事务所名称", "执业证书编号", "组织形式", "注册资本（万元）", "主任会计师/首席合伙人", "批准执业文号", "批准执业日期",
		"分所数量", "注师数量", "联系电话", "传真", "经营场所"}
	err = writeExcel(isExist, firmDetailPath, fireDetailTitle, FirmDetailInfos)
	if err != nil {
		slog.Error("[outPutExcel]", "writeExcel", err)
		return
	}
	return
}

// 写入Excel
func writeExcel(isExist bool, fileName string, titleRow []string, dataSlice interface{}) (err error) {
	sheetName := "Sheet1"
	var file *excelize.File
	if isExist {
		file, err = excelize.OpenFile(fileName)
		if err != nil {
			return
		}
	} else {
		file = excelize.NewFile()
		// 写入自定义标题行
		for col, title := range titleRow {
			cellName := string('A'+col) + "1"
			err = file.SetCellValue(sheetName, cellName, title)
			if err != nil {
				return err
			}
		}
	}

	// 获取已存在的行数
	rows, err := file.GetRows(sheetName)
	if err != nil {
		return err
	}
	startRow := len(rows) + 1
	// 写入结构体数据
	dataSliceValue := reflect.ValueOf(dataSlice)
	if dataSliceValue.Kind() == reflect.Slice {
		sliceLen := dataSliceValue.Len()
		for row := 0; row < sliceLen; row++ {
			data := dataSliceValue.Index(row).Interface().(reptile_model.ExcelWritable)
			cellValues := data.ToExcel()
			for col, value := range cellValues {
				cellName := string('A'+col) + fmt.Sprint(row+startRow)
				err = file.SetCellValue(sheetName, cellName, value)
				if err != nil {
					return err
				}
			}
		}
	}
	// 保存excel文件
	err = file.SaveAs(fileName)
	if err != nil {
		slog.Error("[writeExcel]", "file.SaveAs", err)
		return
	}
	slog.Info("[writeExcel]", "writeExcel", fmt.Sprintf("生成%s成功", fileName))
	return
}

// 读取Excel数据
func readExcel() (infos []reply.ReadExcelReply, err error) {
	// 打开excel文件
	filePath := conf.Static().TemplatePath + "/name.xlsx"
	file, err := xlsx.OpenFile(filePath)
	if err != nil {
		slog.Error("[readExcel]", "xlsx.OpenFile", err)
		return
	}
	// 获取第一个工作表
	sheet := file.Sheets[0]
	skipFirst := true
	// 遍历每一行，跳过第一行
	for _, row := range sheet.Rows {
		if skipFirst {
			skipFirst = false
			continue
		}
		info := reply.ReadExcelReply{}
		for index, cell := range row.Cells {
			switch index {
			case 0:
				info.Id, _ = cell.Int()
			case 1:
				info.Name = cell.String()
			case 2:
				info.No = cell.String()
			case 3:
				info.Address = cell.String()
			}
		}
		infos = append(infos, info)
	}
	return
}

// 获取百度API的token
func getBaiduToken() (token string, err error) {
	uri := "oauth/2.0/token"
	param := map[string]string{
		"client_id":     conf.Baidu().ApiKey,
		"client_secret": conf.Baidu().SecretKey,
		"grant_type":    "client_credentials",
	}
	url := conf.Baidu().Url + uri
	respBody, err := util.HttpGet(url, param, nil)
	var resp reply.AccessTokenResponse
	err = json.Unmarshal(respBody, &resp)
	if err != nil {
		slog.Error("[getBaiduToken]", "json.Unmarshal", err)
		return
	}
	token = resp.AccessToken
	return
}

// 识别验证码文本
func getImageText(req reply.CodeReply) (text string) {
	base64Data := req.VerifyText
	//去掉可能存在的data:image/jpeg;base64,
	dataParts := strings.Split(base64Data, ",")
	if len(dataParts) != 2 {
		slog.Info("[getImageText]", "strings.Split", "invalid base64 data format")
		return
	}
	base64Data = dataParts[1]
	head := map[string]string{
		"Content-Type": "application/x-www-form-urlencoded",
		"Accept":       "application/json",
	}
	param := map[string]string{
		"image":            base64Data,
		"detect_direction": "false",
		"paragraph":        "false",
		"probability":      "false",
	}
	uri := "rest/2.0/ocr/v1/accurate_basic?access_token=" + Token
	url := conf.Baidu().Url + uri
	respBody, err := util.HttpPostByForm(url, param, head)
	if err != nil {
		slog.Error("[getImageText]", "util.HttpPost", err)
		return
	}
	var resp = reply.OCRResult{}
	err = json.Unmarshal(respBody, &resp)
	if err != nil {
		slog.Error("[getImageText]", "json.Unmarshal", err)
		return
	}
	if len(resp.WordsResult) == 0 {
		slog.Error("[getImageText]", "OCR识别错误", "验证长度异常")
		return
	}
	text = strings.ReplaceAll(resp.WordsResult[0].Words, " ", "")
	if len(text) != 4 {
		text = ""
		slog.Error("[getImageText]", "OCR识别错误", "验证长度异常")
		return
	}
	return
}

// 获取政协对应编号
func getAscGuid(address string) (ascGuid string) {
	switch address {
	case "安徽注协":
		return "0000010f-8496-850b-7171-a5a6c127c7e0"
	case "北京注协":
		return "0000010f-8496-8440-e06b-4f9f27a6e22a"
	case "福建注协":
		return "0000010f-8496-851b-06d9-9ce3a3f1c9a7"
	case "上海注协":
		return "0000010f-8496-84dc-eb0d-a1ce842044d0"
	case "江西注协":
		return "0000010f-8496-852a-162e-0bfa034067ce"
	case "江苏注协":
		return "0000010f-8496-84ec-5d56-c3df1c737dc9"
	case "广东注协":
		return "0000010f-8496-8569-ddb2-cd9add2caa43"
	case "河南注协":
		return "0000010f-8496-854a-6dbd-b31b5f4396c9"
	case "湖南注协":
		return "0000010f-8496-8559-4d16-2ae0f99edff7"
	case "浙江注协":
		return "0000010f-8496-84ec-dca5-4437fa1c85f9"
	case "天津注协":
		return "0000010f-8496-847e-921e-7f6839f85c62"
	case "辽宁注协":
		return "0000010f-8496-84bd-ddbd-3b87e41fcb5b"
	case "深圳注协":
		return "0000010f-8496-8598-88bb-d1029f822843"
	case "山东注协":
		return "0000010f-8496-853a-a3b1-01b05b58b16c"
	case "四川注协":
		return "0000010f-8496-85a7-9478-f6a3a445f571"
	case "河北注协":
		return "0000010f-8496-849e-8e6b-bca9192a3ee8"
	case "陕西注协":
		return "0000010f-8496-85e6-49eb-e2bd0daa4450"
	case "湖北注协":
		return "0000010f-8496-854a-60cc-b457629ed137"
	case "重庆注协":
		return "0000010f-8496-85a7-5d79-56dc32768653"
	default:
		return ""
	}
}

// 获取平台验证码
func getPlatformCode(verifyId string) (codeResp reply.CodeReply, err error) {
	// 获取验证码信息
	codeUrl := conf.System().Url + "anon/nvwa-nros/v1/verify-code/get"
	codeParam := map[string]string{
		"verifyId": verifyId,
	}
	codeBytes, err := util.HttpGet(codeUrl, codeParam, nil)
	if err != nil {
		slog.Error("[getPlatformCode]", "util.HTTPGet", err)
		return
	}
	err = json.Unmarshal(codeBytes, &codeResp)
	if err != nil {
		slog.Error("[getPlatformCode]", "json.Unmarshal", err)
		return
	}
	return
}
