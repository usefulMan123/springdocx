package com.example.springdocx.enums;

/**
 * @description:
 * @author: dongkunshuai
 * @date: 2022/2/22 9:56
 */

@SuppressWarnings("all")
public enum ItemEnum {
    HBZJ(1, "货币资金", 1, 2, "货币资金"),
    DQTZ(1, "短期投资", 2, 2, "短期投资"),
    YSPJ(1, "应收票据", 3, 2, "应收票据"),
    YSGL(1,"应收股利", 4,2,"应收股利"),
    YSLX(1, "应收利息", 5, 2, "应收利息"),
    YSZK(1, "应收账款", 6,2, "应收账款"),
    QTYSK(1, "其他应收款", 7, 2, "其他应收款"),
    YFZK(1, "预付账款", 8, 2, "预付账款"),
    YSBTK(1, "应收补贴款", 9, 2, "应收补贴款"),
    CH(1, "存货", 10, 2, "存货"),
    DTFY(1, "待摊费用", 11, 2, "待摊费用"),
    YNNDQDCQZQTZ(1, "一年内到期的长期债券投资", 12, 2, "一年内到期的长期债券投资"),
    QTLDCC(1, "其他流动资产", 13, 2, "其他流动资产"),
    CQGQTZ(1, "长期股权投资", 14, 2, "长期股权投资"),
    CQZCTZ(1, "长期债权投资", 14, 2, "长期债权投资"),
    GDZCYJ(1, "固定资产原价", 15, 2, "固定资产原价"),
    LJZJ(1, "减：累计折旧", 15, 2, "减：累计折旧"),
    GDZCJZZB(1, "减：固定资产减值准备", 15, 2, "减：固定资产减值准备"),
    GCWZ(1, "工程物资", 15,2, "工程物资"),
    ZJGC(1, "在建工程", 15, 2, "在建工程"),
    GDZCQL(1, "固定资产清理", 15, 2, "固定资产清理"),
    WXZC(1, "无形资产", 16, 2, "无形资产"),
    CQDTFY(1, "长期待摊费用", 16, 2, "长期待摊费用"),
    DYZC(1, "递延资产", 16, 2, "递延资产"),
    DYSKJX(1, "递延税款借项", 17, 2, "递延税款借项"),
    DQJK(1, "短期借款", 18, 2, "短期借款"),
    YFPJ(1, "应付票据", 19, 2, "应付票据"),
    YIFZK(1, "应付账款", 20, 2, "应付账款"),
    YUSZK(1, "预收账款", 21 , 2, "预收账款"),
    YFGZ(1, "应付工资", 22, 2, "应付工资"),
    YFFLF(1, "应付福利费", 23, 2, "应付福利费"),
    YFGL(1, "应付股利", 24, 2, "应付股利"),
    YJSJ(1, "应交税金", 25, 2, "应交税金"),
    QTYJK(1, "其他应交款", 26, 2, "其他应交款"),
    QTYFK(1, "其他应付款", 27, 2, "其他应付款"),
    YTFY(1, "预提费用", 28, 2, "预提费用"),
    YJFZ(1, "预计负债", 29, 2, "预计负债"),
    YNNDQCQFZ(1, "一年内到期的长期负债", 30, 2, "一年内到期的长期负债"),
    QTLDFZ(1, "其他流动负债", 31, 2, "其他流动负债"),
    CQJK(1, "长期借款", 32,2, "长期借款"),
    YFZQ(1, "应付债券",32, 2, "应付债券"),
    CQYFK(1, "长期应付款", 32, 2, "长期应付款"),
    ZXYFK(1, "专项应付款", 32, 2, "专项应付款"),
    QTCQFZ(1, "其他长期负债", 32, 2, "其他长期负债"),
    DYSXDX(1, "递延税项贷项", 33, 2, "递延税项贷项"),
    SSZBGB(1, " 实收资本（或股本）", 34, 2, "实收资本（或股本）"),
    YGHTZ(1, "减：已归还投资", 34, 2, "减：已归还投资"),
    SSZBHGBJE(1, "实收资本（或股本）净额", 34, 2, "实收资本（或股本）净额"),
    ZBGJ(1, "资本公积", 35, 2, "资本公积"),
    YYGJ(1, "盈余公积", 36, 2, "盈余公积"),
    BNJLR(2, "本年净利润", 37, 1, "五、净利润（亏损以“－”号填列）"),
    WFPLR(2, "加：年初未分配利润|年末未分配利润", 37, 2, "未分配利润"),
    QTZR(2, "其他转入", 37, 1, "其他转入"),
    ZYYWSR(2, "主营业务收入", 38, 1, "一、主营业务收入"),
    ZYYWCB(2, "主营业务成本", 39, 1, "减：主营业务成本"),
    ZYYWSJJFJ(2, "主营业务税金及附加", 40, 1, "主营业务税金及附加"),
    XSFY(2, "减：销售费用", 41, 1, "减：销售费用"),
    GLFY(2, "管理费用", 42, 1, "管理费用"),
    CWFY(2, "财务费用", 43 , 1, "财务费用"),
    TZSY(2, "投资收益", 44, 1, "加：投资收益（损失以“－”号填列）"),
    BTSR(2, "补贴收入", 45, 1, "补贴收入"),
    YYWSR(2, "营业外收入", 46, 1, "营业外收入"),
    YYWZC(2, "营业外支出", 47, 1, "减：营业外支出"),
    SDSFY(2, "所得税费用", 48, 1, "减：所得税费用"),
    SSGDQY(2, "少数股东权益", 49, 1, "少数股东权益"),
    TQFDYYGJ(2, "提取法定盈余公积", 50, 1, "减：提取法定盈余公积"),
    TQFDGYJ(2, "提取法定公益金", 51, 1, "提取法定公益金"),
    TQCBJJ(2, "提取储备基金",52, 1, "提取储备基金"),
    TQQYFZJJ(2, "提取企业发展基金", 53, 1, "提取企业发展基金"),
    LRGHTZ(2, "利润归还投资", 54, 1, "利润归还投资"),
    YFYXGGL(2, "应付优先股股利", 55, 1, "减：应付优先股股利"),
    TQRYYYGJ(2, "提取任意盈余公积", 56, 1, "提取任意盈余公积"),
    YFPTGGY(2, "应付普通股股利", 57, 1, "应付普通股股利"),
    ZZZBDPTGGL(2, "转作资本（或股本）的普通股股利", 58, 1, "转作资本（或股本）的普通股股利")


    ;


    public static ItemEnum getItem(String cloumn){

        for(ItemEnum item : ItemEnum.values()){

            if(item.getCloumn().equals(cloumn)){

                return item;
            }

        }
        return null;
    }






    //项目类型 1- 年初数-年末数 2- 本年累计
    private Integer type;

    //项目名称
    private String itemName;

    //对应表格的索引
    private Integer index;

    //对应的数据条数
    private Integer count;

    //对应excel 表格数据
    private String cloumn;

    ItemEnum(Integer type, String itemName, Integer index, Integer count, String cloumn) {
        this.type = type;
        this.itemName = itemName;
        this.index = index;
        this.count = count;
        this.cloumn = cloumn;
    }

    public String getCloumn() {
        return cloumn;
    }

    public void setCloumn(String cloumn) {
        this.cloumn = cloumn;
    }

    public Integer getType() {
        return type;
    }

    public void setType(Integer type) {
        this.type = type;
    }

    public String getItemName() {
        return itemName;
    }

    public void setItemName(String itemName) {
        this.itemName = itemName;
    }

    public Integer getIndex() {
        return index;
    }

    public void setIndex(Integer index) {
        this.index = index;
    }

    public Integer getCount() {
        return count;
    }

    public void setCount(Integer count) {
        this.count = count;
    }

    ItemEnum() {
    }

    ItemEnum(Integer type, String itemName, Integer index, Integer count) {
        this.type = type;
        this.itemName = itemName;
        this.index = index;
        this.count = count;
    }

}
