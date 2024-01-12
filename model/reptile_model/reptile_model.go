package reptile_model

// ExcelWritable 定义 ExcelWritable 接口
type ExcelWritable interface {
	ToExcel() []string
}

// OfficeDetail 事务所数据
type OfficeDetail struct {
	OffName       string `json:"offName"`       // 会计师事务所名称
	SubCount      string `json:"subCount"`      // 分所数量
	PartnerCount  string `json:"partnerCount"`  // 合伙人或股东人数
	CpaNum        string `json:"cpaNum"`        // 注册会计师人数
	OndutySum     string `json:"ondutySum"`     // 从业人员人数
	AllPerCount   string `json:"allPerCount"`   // 注册会计师人数 （含分所）
	NoAllPerCount string `json:"noAllPerCount"` // 从业人员人数 （含分所）
}

// FirmDetail 事务所详情数据
type FirmDetail struct {
	OffName     string `json:"offName"`     // 事务所名称
	OffCode     string `json:"offCode"`     // 执业证书编号
	OffType     string `json:"offType"`     // 组织形式
	RegMoney    string `json:"regMoney"`    // 注册资本（万元）
	AccountName string `json:"accountName"` // 主任会计师/首席合伙人
	PassWord    string `json:"passWord"`    // 批准执业文号
	PassTime    string `json:"passTime"`    // 批准执业日期
	SubCount    string `json:"subCount"`    // 分所数量
	CpaNum      string `json:"cpaNum"`      // 注师数量
	PhoneDecode string `json:"phoneDecode"` // 联系电话
	Fax         string `json:"fax"`         // 传真
	OfficeAddr  string `json:"officeAddr"`  // 经营场所
}

// PartnerInfo 合伙人数据
type PartnerInfo struct {
	OffName string `json:"OFF_NAME"` // 会计师事务所名称
	Number  string `json:"Number"`   // 序号
	PerName string `json:"PER_NAME"` // 合伙人（股东姓名）
	IsCPA   string `json:"IS_CPA"`   // 是否注师（0：否，1：是）
	PerCode string `json:"PER_CODE"` // 注师编号
}

// AccountantInfo 注册会计师数据
type AccountantInfo struct {
	OffName string `json:"offName"` // 会计师事务所名称
	Number  string `json:"Number"`  // 序号
	PerName string `json:"perName"` // 姓名
	PerCode string `json:"perCode"` // 人员编号
	Gender  string `json:"gender"`  // 性别
	RegWord string `json:"regWord"` // 批准文号
}

// PractitionerInfo 从业人员数据
type PractitionerInfo struct {
	OffName  string `json:"OFF_NAME"`  // 会计师事务所名称
	Number   string `json:"Number"`    // 序号
	EmpName  string `json:"EMP_NAME"`  // 姓名
	Gender   string `json:"GENDER"`    // 性别（1：男，2：女）
	IntoTime string `json:"INTO_TIME"` // 进所时间
	IsPact   string `json:"IS_PACT"`   // 是否签合同（1：是，2：否）
	IsSafety string `json:"IS_SAFETY"` // 是否参加社保（1：是，2：否）
	IsCpm    string `json:"IS_CPM"`    // 是否党员（1：是，2：否）
}

// FirmBranchInfo 事务分所人数数据
type FirmBranchInfo struct {
	OffName   string `json:"offName"`   // 会计师事务所名称
	CpaNum    string `json:"cpaNum"`    // 注册会计师人数
	OndutySum string `json:"ondutySum"` // 从业人员人数
}

// ToExcel 实现 ExcelWritable 接口的 ToExcel 方法
func (o OfficeDetail) ToExcel() []string {
	return []string{o.OffName, o.SubCount, o.PartnerCount, o.CpaNum, o.OndutySum, o.AllPerCount, o.NoAllPerCount}
}

// ToExcel 实现 ExcelWritable 接口的 ToExcel 方法
func (o PartnerInfo) ToExcel() []string {
	return []string{o.OffName, o.Number, o.PerName, o.IsCPA, o.PerCode}
}

// ToExcel 实现 ExcelWritable 接口的 ToExcel 方法
func (o AccountantInfo) ToExcel() []string {
	return []string{o.OffName, o.Number, o.PerName, o.PerCode, o.Gender, o.RegWord}
}

// ToExcel 实现 ExcelWritable 接口的 ToExcel 方法
func (o PractitionerInfo) ToExcel() []string {
	return []string{o.OffName, o.Number, o.EmpName, o.Gender, o.IntoTime, o.IsPact, o.IsSafety, o.IsCpm}
}

// ToExcel 实现 ExcelWritable 接口的 ToExcel 方法
func (o FirmBranchInfo) ToExcel() []string {
	return []string{o.OffName, o.CpaNum, o.OndutySum}
}

// ToExcel 实现 ExcelWritable 接口的 ToExcel 方法
func (o FirmDetail) ToExcel() []string {
	return []string{o.OffName, o.OffCode, o.OffType, o.RegMoney, o.AccountName, o.PassWord, o.PassTime, o.SubCount, o.CpaNum,
		o.PhoneDecode, o.Fax, o.OfficeAddr}
}
