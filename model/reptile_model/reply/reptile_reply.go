package reply

import "reptile/model/reptile_model"

type ReadExcelReply struct {
	Id      int    `json:"id"`      // 序号
	Name    string `json:"name"`    // 会计师事务所名称
	No      string `json:"no"`      // 执业证书编号
	Address string `json:"address"` // 地区
}

type CodeReply struct {
	VerifyText string `json:"verifyText"` // 验证码图片链接
	VerifyId   string `json:"verifyId"`   // 验证码图片ID
}

// AccessTokenResponse 百度API请求Token返回值
type AccessTokenResponse struct {
	RefreshToken  string `json:"refresh_token"`
	ExpiresIn     int    `json:"expires_in"`
	SessionKey    string `json:"session_key"`
	AccessToken   string `json:"access_token"`
	Scope         string `json:"scope"`
	SessionSecret string `json:"session_secret"`
}

// OCRResult 百度API图片识别返回
type OCRResult struct {
	WordsResult    []WordResult `json:"words_result"`
	WordsResultNum int          `json:"words_result_num"`
	LogID          int64        `json:"log_id"`
}

type WordResult struct {
	Words string `json:"words"`
}

// FirmBranchInfo 事务所分所数据
type FirmBranchInfo struct {
	OffCode string `json:"offCode"`
	OffName string `json:"OffName"`
}

type Firm struct {
	OffName       string `json:"offName"`       // 会计师事务所名称
	SubCount      string `json:"subCount"`      // 分所数量
	PartnerCount  string `json:"partnerCount"`  // 合伙人或股东人数
	CpaNum        string `json:"cpaNum"`        // 注册会计师人数
	OndutySum     string `json:"ondutySum"`     // 从业人员人数
	AllPerCount   string `json:"allPerCount"`   // 注册会计师人数 （含分所）
	NoAllPerCount string `json:"noAllPerCount"` // 从业人员人数 （含分所）
	OffCode       string `json:"offCode"`
	OffType       string `json:"offType"`
	RegMoney      string `json:"regMoney"`
	AccountName   string `json:"accountName"`
	PassWord      string `json:"passWord"`
	PassTime      string `json:"passTime"`
	PhoneDecode   string `json:"phoneDecode"`
	Fax           string `json:"fax"`
	OfficeAddr    string `json:"officeAddr"`
}

// OfficeInfo 会计事务所信息接口返回
type OfficeInfo struct {
	Result struct {
		List []struct {
			Id         string `json:"id"`
			OffAllCode string `json:"offAllcode"`
		} `json:"list"`
	} `json:"info"`
}

type RegisterInfoReply struct {
	Info struct {
		SubOfficeList []FirmBranchInfo `json:"subOfficeList"`
		HeadInfo      Firm             `json:"headInfo"`
	} `json:"info"`
}

type PartnerInfoReply struct {
	Info struct {
		OffName   string `json:"OFF_NAME"`
		ResultMap struct {
			Rows []reptile_model.PartnerInfo `json:"rows"`
		} `json:"resultMap"`
	} `json:"info"`
}

type AccountantInfoReply struct {
	Info struct {
		List []reptile_model.AccountantInfo `json:"list"`
	} `json:"info"`
}

type PractitionerInfoReply struct {
	Info struct {
		OffName   string `json:"OFF_NAME"`
		ResultMap struct {
			Rows []reptile_model.PractitionerInfo `json:"rows"`
		} `json:"resultMap"`
	} `json:"Info"`
}

type FirmBranchInfoReply struct {
	Info struct {
		HeadInfo reptile_model.FirmBranchInfo `json:"headInfo"`
	} `json:"info"`
}

// BranchReply 分所所有数据返回
type BranchReply struct {
	FirmBranchInfos        []reptile_model.FirmBranchInfo   `json:"FirmBranchInfos"`
	BranchAccountantInfos  []reptile_model.AccountantInfo   `json:"BranchAccountantInfos"`
	BranchPractitionerInfo []reptile_model.PractitionerInfo `json:"BranchPractitionerInfo"`
}
