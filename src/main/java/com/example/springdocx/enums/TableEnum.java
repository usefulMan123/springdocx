package com.example.springdocx.enums;

/**
 * 维护项目信息
 */
@SuppressWarnings("all")
public enum TableEnum {

    HBZJ(1, "货币资金", 1),
    DQTZ(2, "短期投资", 1),
    YSPJ(3, "应收票据", 1),
    YSGL(4, "应收股利", 1),
    YSLX(5, "应收利息",1),
    YSZK(6, "应收账款",1),
    QTYSK(7, "其他应收款",1),
    YFZK(8, "预付账款", 1),
    YSBTK(9, "应收补贴款", 1),
    CH(10, "存货", 1),
    DTFY(11, "待摊费用", 1),
    YNNDQDCQZQTZ(12, "一年内到期的长期债券投资", 1),
    QTLDZC(13, "其他流动资产", 1),
    CQTZ(14, "长期投资", 1),
    GDZC(15, "固定资产", 1),
    WXZCJQTZC(16, "无形资产及其他资产",1),
    DYSX(17, "递延税项", 1),
    DQJK(18, "短期借款", 1),
    YFPJ(19, "应付票据", 1),
    YIFZK(20, "应付账款",1),
    YUSZK(21, "预收账款", 1),
    YFGZ(22, "应付工资",1),
    YFFLF(23, "应付福利费", 1),
    YFGL(24, "应付股利",1),
    YJSJ(25, "应交税金",1),
    QTYJK(26, "其他应交款",1),
    QTYFK(27, "其他应付款", 1),
    YTFY(28, "预提费用",1),
    YJFZ(29, "预计费用",1),
    YNNDQCQFZ(30, "一年内到期的长期负债",1),
    QTLDFZ(31, "其他流动负债",1),
    CQFZ(32, "长期负债", 1),
    DYSXDX(33, "递延税项", 1),
    SSZB(34, "实收资本", 0),
    ZBGJ(35, "资本公积", 1),
    YYGJ(36, "盈余公积",1),
    WFPLR(37, "未分配利润", 0),
    ZYYWCB(38, "主营业务成本", 1),
    ZYYWSR(39, "主营业务收入", 1),
    ZYYWSJJFJ(40, "主营业务税金及附加", 1),
    XSFY(41, "销售费用", 1),
    GLFY(42, "管理费用", 1),
    CWFY(43, "财务费用", 1),
    TZSY(44, "投资收益", 1),
    BTSR(45, "补贴收入", 1),
    YYWSR(46, "营业外收入", 1),
    YYWZC(47, "营业外支出", 1),
    SDSFY(48, "所得税费用",1),
    SSGDQY(49, "少数股东权益",1),
    TQFDYYGJ(50, "提取法定公积盈余",1),
    TQFDGYJ(51, "提取法定公益金",1),
    TQCBJJ(52, "提取储备基金", 1),
    TQQYFZJJ(53, "提取企业发展基金", 1),
    LRGHTZ(54, "利润归还投资", 1),
    YFYXGGL(55,"应付优先股股利", 1),
    TQRYYYGJ(56, "提取任意盈余公积", 1),
    YFPTGGL(57, "应付普通股股利", 1),
    ZZZBPTGGL(58, "转作资本（或股本）的普通股股利", 1)

    ;

    //获取表格 标题信息
    public static String getTitle(Integer index){

        for(TableEnum table : TableEnum.values()){

            if(table.getIndex() == index){
                return table.getTitle();
            }

        }
        return null;
    }

    public static Integer getIsHj(Integer index){
        for(TableEnum table : TableEnum.values()){

            if(table.getIndex() == index){
                return table.getIsHj();
            }

        }
        return null;
    }

    // 表格索引，唯一值
    private Integer index;

    // 表格标题
    private String title;

    //是否需要合计
    private Integer isHj;



    public void setIndex(Integer index) {
        this.index = index;
    }

    public Integer getIsHj() {
        return isHj;
    }

    public void setIsHj(Integer isHj) {
        this.isHj = isHj;
    }

    TableEnum(Integer index, String title, Integer isHj) {
        this.index = index;
        this.title = title;
        this.isHj = isHj;
    }

    TableEnum(int index, String title) {
        this.index = index;
        this.title = title;
    }

    TableEnum() {
    }

    public int getIndex() {
        return index;
    }

    public void setIndex(int index) {
        this.index = index;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }
}
