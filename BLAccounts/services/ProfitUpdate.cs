using System;
using System.Data;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;


using DataBase;
using DataBase_Oracle.Connections;
using XL.XSheet;

namespace BLAccounts

{
    public class ProfitUpdateService : BL_Base
    {

        string type = "";
        string company_code = "";
        string branch_code = "";
        string year_code = "";
        string ErrorMessage = "";
        

        string sql1 = "";

        public IDictionary<string, object> ProcessProfit(Dictionary<string, object> SearchData)
        {

            Dictionary<string, object> RetData = new Dictionary<string, object>();
            
            ErrorMessage = "";
            try
            {

                type = SearchData["type"].ToString();
                company_code = SearchData["company_code"].ToString();
                year_code = SearchData["year_code"].ToString();
                


                if (ErrorMessage != "")
                {
                    if (Con_Oracle != null)
                        Con_Oracle.CloseConnection();
                    throw new Exception(ErrorMessage);
                }

                if (type == "AIR-EXPORT-FORWARDING")
                {
                    sql = " insert into profitreport( ";
                    sql += " type, jv_year, mblid, reportdate, mbl_date, mblslno, branch, mblno, sino, hblno, buy_date, sell_date, exporter_name, consignee_name, agent_name, sman_name, nomination, ";
                    sql += " mbl_frt_status, hbl_frt_status,pkg,pcs, mbl_chwt, hbl_chwt, mbl_grwt, hbl_grwt, pol, pod, pofd, fin_year, folder_no, inv_nos, shpr_location, ";
                    sql += " shpr_state, shpr_created, notify, liner, commodity, pod_country, buyer_name, ";
                    sql += " ex_01, ex_02, ex_03, ex_17, ex_ot, ";
                    sql += " in_01, in_02, in_03, in_17, in_ot, in_10, ";
                    sql += " cost_dr, cost_cr, rebate_dr, buy, sell ";
                    sql += " ) ";
                    sql += " select 'AIR-EXPORT-FWD', " + year_code +",  hbl_pkid, hbl_date, mbl_date,mblslno, branch,";
                    sql += " mblno, SINO,  hblno,";
                    sql += " buy_date,sell_date,";
                    sql += " exp.cust_name as exporter_name,";
                    sql += " imp.cust_name as consignee_name,";
                    sql += " agent.cust_name as agent_name,";
                    sql += " nvl(sman1.param_name,nvl(sman2.param_name,sman.param_name)) as sman_name,";
                    sql += " nvl(hbl_nomination,imp.cust_nomination) as nomination,";
                    sql += " mbl_frt_status,";
                    sql += " hbl_frt_status,";
                    sql += " hbl_pkg,";
                    sql += " hbl_pcs,";
                    sql += " mbl_chwt,";
                    sql += " hbl_chwt,";
                    sql += " mbl_grwt,";
                    sql += " hbl_grwt,";
                    sql += " pol.param_name as pol, ";
                    sql += " pod.param_name as pod, ";
                    sql += " pofd.param_name as pofd,";
                    sql += " fin_year,";
                    sql += " folder_no, ";
                    sql += " inv_nos, ";
                    sql += " expadd.add_city as shpr_location, ";
                    sql += " expstate.param_name as shpr_state, ";
                    sql += " exp.rec_created_date as exp_created,";
                    sql += " notify.bl_notify_name as notify, ";
                    sql += " liner.param_name as airline, ";
                    sql += " cmdty.param_name as commodity,";
                    sql += " podcntry.param_name as pod_country, ";
                    sql += " buyer.cust_name as buyer_name,";
                    sql += " frt_dr,";
                    sql += " fsc_dr,";
                    sql += " wrs_dr,";
                    sql += " mcc_dr,";
                    sql += " oth_dr,";
                    sql += " frt_cr,";
                    sql += " fsc_cr,";
                    sql += " wrs_cr,";
                    sql += " mcc_cr,";
                    sql += " oth_cr,";
                    sql += " margin_cr,";
                    sql += " frt_ho_dr,";
                    sql += " frt_ho_cr,";
                    sql += " rebate_dr,";
                    sql += " buy,";
                    sql += " sell";
                    sql += " from (";
                    sql += " select mbl.hbl_pkid, mbl.hbl_date, mbl.hbl_date as mbl_date,mbl.hbl_no as mblslno, a.rec_branch_code as branch,";
                    sql += " mbl.hbl_bl_no as mblno, hbl.hbl_no as SINO,  hbl.hbl_bl_no as hblno,";
                    sql += " max(case when jvh_type = 'PN' then jvh_date else null end ) as buy_date,";
                    sql += " max(case when jvh_type = 'IN' then jvh_date else null end ) as sell_date,";
                    sql += " max(mbl.hbl_terms) as mbl_frt_status,";
                    sql += " max(hbl.hbl_terms) as hbl_frt_status,";

                    sql += " max(hbl.hbl_pkg) as hbl_pkg,";
                    sql += " max(hbl.hbl_pcs) as hbl_pcs,";

                    sql += " max(mbl.hbl_chwt) as mbl_chwt,";
                    sql += " max(hbl.hbl_chwt) as hbl_chwt,";
                    sql += " max(mbl.hbl_grwt) as mbl_grwt,";
                    sql += " max(hbl.hbl_grwt) as hbl_grwt,";
                    sql += " max(a.jvh_year) as fin_year,";
                    sql += " max(mbl.hbl_folder_no) as folder_no, ";
                    sql += " max(hbl.hbl_ar_invnos) as inv_nos, ";
                    sql += " max(hbl.hbl_exp_id) as hbl_ex_id,";
                    sql += " max(hbl.rec_branch_code) as branch_code, ";
                    sql += " max(hbl.hbl_exp_id) as hbl_exp_id, ";
                    sql += " max(hbl.hbl_salesman_id) as hbl_salesman_id,";
                    sql += " max(hbl.hbl_nomination) as hbl_nomination,";
                    sql += " max(hbl.hbl_imp_id) as hbl_imp_id,";
                    sql += " max(hbl.hbl_agent_id) as hbl_agent_id,";
                    sql += " max(hbl.hbl_pol_id) as hbl_pol_id,";
                    sql += " max(hbl.hbl_pod_id) hbl_pod_id,";
                    sql += " max(hbl.hbl_pofd_id) as hbl_pofd_id,";
                    sql += " max(hbl.hbl_exp_br_id) as hbl_exp_br_id,";
                    sql += " max(hbl.hbl_carrier_id) as hbl_carrier_id,";
                    sql += " max(hbl.hbl_commodity_id) as hbl_commodity_id,";
                    sql += " max(hbl.hbl_pod_country_id) as hbl_pod_country_id, ";
                    sql += " max(hbl.hbl_buyer_id) as hbl_buyer_id,";
                    sql += " sum( case when jv_drcr = 'DR' and acc_code  in('1205001') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	frt_dr,";
                    sql += " sum( case when jv_drcr = 'DR' and acc_code  in('1205002') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	fsc_dr,";
                    sql += " sum( case when jv_drcr = 'DR' and acc_code  in('1205003') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	wrs_dr,";
                    sql += " sum( case when jv_drcr = 'DR' and acc_code  in('1205017') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	mcc_dr,";
                    sql += " sum( case when jv_drcr = 'DR' and acc_code  like '1205%'  and acc_code not in ('1205001','1205002', '1205003', '1205017') and  jvh_type not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	oth_dr,";
                    sql += " sum( case when jv_drcr = 'CR' and acc_code  in('1205001') and jvh_type  not in('HO','IN-ES')   then ABS(ct_amount) else 0 end ) as 	frt_cr,";
                    sql += " sum( case when jv_drcr = 'CR' and acc_code  in('1205002') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	fsc_cr,";
                    sql += " sum( case when jv_drcr = 'CR' and acc_code  in('1205003') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	wrs_cr,";
                    sql += " sum( case when jv_drcr = 'CR' and acc_code  in('1205017') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	mcc_cr,";
                    sql += " sum( case when jv_drcr = 'CR' and acc_code  like '1205%'  and acc_code not in ('1205001','1205002', '1205003', '1205017',  '1205010') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	oth_cr,";
                    sql += " sum( case when jv_drcr = 'CR' and acc_code  in('1205010') and jvh_type  not in('HO','IN-ES')  then ct_amount else 0 end ) as 	margin_cr,";
                    sql += " sum( case when jvh_type  in('HO','IN-ES')  and jv_drcr = 'DR' and acc_code  in('1205001')  then ABS(ct_amount) else 0 end ) as 	frt_ho_dr,";
                    sql += " sum( case when jvh_type  in('HO','IN-ES')  and jv_drcr = 'CR' and acc_code  in('1205001')  then ABS(ct_amount) else 0 end ) as 	frt_ho_cr,";
                    sql += " sum( case when jv_drcr = 'DR' and acc_code  in('1205100')  then ABS(ct_amount) else 0 end ) as 	rebate_dr,";
                    sql += " sum( case when jv_drcr = 'DR' and acc_code  like '1205%'  then ABS(ct_amount) else 0 end ) as 	buy,";
                    sql += " sum( case when jv_drcr = 'CR' and acc_code  like '1205%'  then ABS(ct_amount) else 0 end ) as 	sell";
                    sql += " from ledgerh a";
                    sql += " inner join ledgert b on  jvh_pkid = jv_parent_id ";
                    sql += " inner join costcentert c on b.jv_pkid = c.ct_jv_id";
                    sql += " inner join hblm hbl on c.ct_cost_id = hbl.hbl_pkid";
                    sql += " inner join hblm mbl on hbl.hbl_mbl_id = mbl.hbl_pkid";
                    sql += " inner join acctm e on jv_acc_id = acc_pkid";
                    sql += " where a.jvh_year = " +  year_code  + " and  a.rec_company_code = '{COMPCODE}' ";
                    //sql += " and a.rec_branch_code = '{BRCODE}' ";
                    sql += " and hbl.hbl_type = 'HBL-AE'  ";
                    sql += " group by a.rec_branch_code,mbl.hbl_pkid, mbl.hbl_no, mbl.hbl_bl_no, mbl.hbl_date, hbl.hbl_pkid, hbl.hbl_no, hbl.hbl_bl_no";
                    sql += " )a ";
                    sql += " left join customerm exp on hbl_exp_id = exp.cust_pkid";
                    sql += " left join custdet  cd on branch_code = cd.det_branch_code and hbl_exp_id = cd.det_cust_id ";
                    sql += " left join param sman1 on hbl_salesman_id = sman1.param_pkid";
                    sql += " left join param sman on exp.cust_sman_id = sman.param_pkid";
                    sql += " left join param sman2 on cd.det_sman_id = sman2.param_pkid";
                    sql += " left join customerm imp on hbl_imp_id = imp.cust_pkid";
                    sql += " left join customerm agent on hbl_agent_id = agent.cust_pkid";
                    sql += " left join param pol on hbl_pol_id = pol.param_pkid";
                    sql += " left join param pod on hbl_pod_id = pod.param_pkid";
                    sql += " left join param pofd on hbl_pofd_id = pofd.param_pkid";
                    sql += " left join addressm expadd on hbl_exp_br_id = expadd.add_pkid";
                    sql += " left join param expstate on expadd.add_state_id = expstate.param_pkid";
                    sql += " left join bl notify on hbl_pkid = notify.bl_pkid";
                    sql += " left join param liner on hbl_carrier_id = liner.param_pkid";
                    sql += " left join param cmdty on hbl_commodity_id = cmdty.param_pkid";
                    sql += " left join param podcntry on hbl_pod_country_id = podcntry.param_pkid ";
                    sql += " left join customerm buyer on hbl_buyer_id = buyer.cust_pkid ";

                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{COMPCODE}", company_code);


                    Con_Oracle = new DBConnection();

                    Con_Oracle.BeginTransaction();

                    sql1 = " delete from profitreport where type = 'AIR-EXPORT-FWD' and jv_year =" + year_code;
                    Con_Oracle.ExecuteNonQuery(sql1);

                    Con_Oracle.ExecuteNonQuery(sql);

                    sql = " update profitreport set profit = nvl(sell,0) - nvl(buy,0) where type ='AIR-EXPORT-FWD'";
                    Con_Oracle.ExecuteNonQuery(sql);
                    sql = " update profitreport set roi = profit / buy *100 where buy > 0 and type ='AIR-EXPORT-FWD'";
                    Con_Oracle.ExecuteNonQuery(sql);

                    /*
                    sql = " update profitreport p1 set profit = ";
                    sql += " ( select   sum(nvl(p2.sell, 0) - nvl(p2.buy, 0)) * p1.hbl_chwt / sum(p2.hbl_chwt) ";
                    sql += " from profitreport p2 ";
                    sql += "  where p1.mblid = p2.mblid )";
                    Con_Oracle.ExecuteNonQuery(sql);
                    */


                    Con_Oracle.CommitTransaction();

                    Con_Oracle.CloseConnection();
                }

                if (type == "SEA-EXPORT-FORWARDING")
                {

                    sql += " insert into profitreport( ";
                    sql += " type, jv_year, mblid,reportdate,branch,mblno,mbl_date,mblslno,mbl_status,folder_no ";
                    sql += " ,job_type,sino,hblno ";
                    sql += " ,hbl_date,buy_date,sell_date ";
                    sql += " ,exporter_name,consignee_name,agent_name,sman_name ";
                    sql += " ,nomination,mbl_frt_status,hbl_frt_status ";
                    sql += " ,pkg,pcs,hbl_grwt,hbl_cbm ";
                    sql += " ,pol,pod,pofd,etd,cntr,teu ";
                    sql += " ,fin_year,inv_nos ";
                    sql += " ,shpr_location,shpr_state,shpr_created ";
                    sql += " ,notify,ex_works,ddp,ddu,liner,commodity ";
                    sql += " ,pod_country,buyer_name ";
                    sql += " ,ex_05,ex_06,ex_07 ";
                    sql += " ,in_05,in_06,in_07 ";
                    sql += " ,cost_dr,cost_cr,rebate_dr,buy,sell ";
                    sql += " )";

                    sql += " select 'SEA-EXPORT-FWD', " +  year_code +",  mblid,";
                    sql += "mbl_date as ddate,";
                    sql += " branch, ";
                    sql += " mblno, ";
                    sql += " mbl_date, ";
                    sql += " mblslno,";
                    sql += " status.param_name as mbl_status,";
                    sql += " mbl_folder_no,";
                    sql += " shipment_type, ";
                    sql += " SINO,  ";
                    sql += " hblno, ";
                    sql += " hbl_date, ";
                    sql += " buy_date, ";
                    sql += " sell_date,";
                    sql += " exp.cust_name as exporter_name,";
                    sql += " imp.cust_name as consignee_name, ";
                    sql += " agent.cust_name as agent_name,";
                    sql += " nvl(sman1.param_name,nvl(sman2.param_name,sman.param_name)) as sman_name, ";
                    sql += " nvl(hbl_nomination,imp.cust_nomination) as nomination,";
                    sql += " mbl_frt_status,";
                    sql += " hbl_frt_status,";

                    sql += " hbl_pkg,";
                    sql += " hbl_pcs,";

                    sql += " hbl_grwt,";
                    sql += " hbl_cbm,";
                    sql += " pol.param_name as pol,";
                    sql += " pod.param_name as pod,";
                    sql += " pofd.param_name as pofd,";
                    sql += " etd,";
                    sql += " cntr,";
                    sql += " teu, ";
                    sql += " fin_year, ";
                    sql += " inv_nos,  ";
                    sql += " expadd.add_city as shpr_location,";
                    sql += " expstate.param_name as shpr_state,  ";
                    sql += " exp.rec_created_date as shpr_created, ";
                    sql += " notify.bl_notify_name as notify, ";
                    sql += " ex_works,";
                    sql += " ddp,  ";
                    sql += " ddu,";
                    sql += " liner.param_name as liner, ";


                    sql += " cmdty.param_name as commodity,  ";
                    sql += " podcntry.param_name as pod_country,";
                    sql += " buyer.cust_name as buyer_name,  ";
                    sql += " ex_1105, ";
                    sql += " ex_1106,";
                    sql += " ex_1107,";
                    sql += " in_1105,";
                    sql += " in_1106, ";
                    sql += " in_1107,";
                    sql += " cost_dr,";
                    sql += " cost_cr,";
                    sql += " rebate_dr,";
                    sql += " buy, sell";

                    sql += "  from ( ";


                    sql += " select mbl.hbl_pkid as mblid, mbl.hbl_date as mbl_date, mbl.hbl_no as mblslno,a.rec_branch_code as branch, ";
                    sql += " mbl.hbl_bl_no as mblno, hbl.hbl_no as SINO,  hbl.hbl_bl_no as hblno,";
                    sql += " hbl.hbl_pkid, ";
                    sql += " max(case when jvh_type = 'PN' then jvh_date else null end ) as buy_date,";
                    sql += " max(case when jvh_type = 'IN' then jvh_date else null end ) as sell_date,";
                    sql += " max(hbl.hbl_exp_id) as hbl_exp_id, ";
                    sql += " max(hbl.hbl_exp_br_id) as hbl_exp_br_id, ";
                    sql += " max(hbl.hbl_imp_id) as hbl_imp_id, ";
                    sql += " max(hbl.hbl_agent_id) as hbl_agent_id,";
                    sql += " max(hbl.rec_branch_code) as hbl_rec_branch_code, ";
                    sql += " max(mbl.hbl_terms) as mbl_frt_status,";
                    sql += " max(hbl.hbl_terms) as hbl_frt_status,";
                    sql += " max(hbl.hbl_cbm) as hbl_cbm,";
                    sql += " max(hbl.hbl_salesman_id) as hbl_salesman_id,";
                    sql += " max(hbl.hbl_nomination) as hbl_nomination,";

                    sql += " max(hbl.hbl_pkg) as hbl_pkg,";
                    sql += " max(hbl.hbl_pcs) as hbl_pcs,";

                    sql += " max(hbl.hbl_grwt) as hbl_grwt,";
                    sql += " max(hbl.hbl_pol_id) as hbl_pol_id,";
                    sql += " max(hbl.hbl_pod_id) as hbl_pod_id,";
                    sql += " max(hbl.hbl_pofd_id) as hbl_pofd_id,";
                    sql += " max(hbl.hbl_pod_country_id) as hbl_pod_country_id,";
                    sql += " max(hbl.hbl_buyer_id) as hbl_buyer_id,  ";
                    sql += " max(a.jvh_year) as fin_year, ";
                    sql += " max(mbl.hbl_folder_no) as mbl_folder_no,";
                    sql += " max(mbl.hbl_pol_etd) as etd,";
                    sql += " max(hbl.hbl_book_cntr) as cntr,";
                    sql += " max(mbl.hbl_status_id) as mbl_status_id,";
                    sql += " max(hbl.hbl_book_cntr_teu) as teu,  ";
                    sql += " max(hbl.hbl_date) as hbl_date, ";
                    sql += " max(hbl.hbl_ar_invnos) as inv_nos,  ";
                    sql += " max(hbl.hbl_carrier_id) as hbl_carrier_id, ";
                    sql += " max(hbl.hbl_commodity_id) as hbl_commodity_id,  ";
                    sql += " max(hbl.hbl_ddp) as ddp,  ";
                    sql += " max(hbl.hbl_ddu) as ddu,";
                    sql += " max(hbl.hbl_ex_works) as ex_works,";
                    sql += " max(mbl.hbl_shipment_type) as shipment_type, ";
                    sql += " sum( case when jv_drcr = 'DR' and acc_main_code  in('1105') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as ex_1105,";
                    sql += " sum( case when jv_drcr = 'DR' and acc_main_code  in('1106') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as ex_1106, ";
                    sql += " sum( case when jv_drcr = 'DR' and acc_main_code  in('1107') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as ex_1107, ";
                    sql += " sum( case when jv_drcr = 'CR' and acc_main_code  in('1105') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as in_1105,";
                    sql += " sum( case when jv_drcr = 'CR' and acc_main_code  in('1106') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as in_1106, ";
                    sql += " sum( case when jv_drcr = 'CR' and acc_main_code  in('1107') and jvh_type  not in('HO','IN-ES')   then ABS(ct_amount) else 0 end ) as in_1107,";
                    sql += " sum( case when jvh_type  in('HO','IN-ES')  and jv_drcr = 'DR' and acc_main_code  in('1105')  then ABS(ct_amount) else 0 end ) as  cost_dr, ";
                    sql += " sum( case when jvh_type  in('HO','IN-ES')  and jv_drcr = 'CR' and acc_main_code  in('1105')  then ABS(ct_amount) else 0 end ) as cost_cr,";
                    sql += " sum( case when jv_drcr = 'DR' and acc_code  in('1105100','1106100','1107100')  then ABS(ct_amount) else 0 end ) as rebate_dr,";
                    sql += " sum( case when jv_drcr = 'DR' and acc_main_code in('1105','1106','1107') then ABS(ct_amount) else 0 end ) as buy, ";
                    sql += " sum( case when jv_drcr = 'CR' and acc_main_code in('1105','1106','1107')    then ABS(ct_amount) else 0 end ) as sell     ";
                    sql += " from ledgerh a ";
                    sql += " inner join ledgert b on  jvh_pkid = jv_parent_id   ";
                    sql += " inner join costcentert c on b.jv_pkid = c.ct_jv_id  ";
                    sql += " inner join hblm hbl on c.ct_cost_id = hbl.hbl_pkid  ";
                    sql += " inner join hblm mbl on hbl.hbl_mbl_id = mbl.hbl_pkid  ";
                    sql += " inner join acctm e on jv_acc_id = acc_pkid  ";
                    sql += "  where a.jvh_year = " + year_code +  " and a.rec_company_code = '{COMPCODE}' ";

                   //sql += " and a.rec_branch_code = '{BRCODE}' ";

                    sql += " and hbl.hbl_type = 'HBL-SE' ";
                    sql += " group by a.rec_branch_code,mbl.hbl_pkid, mbl.hbl_no, mbl.hbl_bl_no, mbl.hbl_date, hbl.hbl_pkid, hbl.hbl_no, hbl.hbl_bl_no  ";
                    sql += " ) a ";
                    sql += " left join customerm exp on a.hbl_exp_id = exp.cust_pkid  ";
                    sql += " left join custdet  cd on a.hbl_rec_branch_code = cd.det_branch_code and a.hbl_exp_id = cd.det_cust_id   ";
                    sql += " left join param sman on exp.cust_sman_id = sman.param_pkid  ";
                    sql += " left join param sman2 on cd.det_sman_id = sman2.param_pkid  ";
                    sql += " left join param sman1 on a.hbl_salesman_id = sman1.param_pkid";

                    sql += " left join customerm imp on a.hbl_imp_id = imp.cust_pkid  ";
                    sql += " left join customerm agent on a.hbl_agent_id = agent.cust_pkid ";
                    sql += " left join param pol on a.hbl_pol_id = pol.param_pkid ";
                    sql += " left join param pod on a.hbl_pod_id = pod.param_pkid ";
                    sql += " left join param pofd on a.hbl_pofd_id = pofd.param_pkid ";
                    sql += " left join param podcntry on a.hbl_pod_country_id = podcntry.param_pkid  ";
                    sql += " left join customerm buyer on a.hbl_buyer_id = buyer.cust_pkid  ";
                    sql += " left join param status on a.mbl_status_id = status.param_pkid ";
                    sql += " left join addressm expadd on a.hbl_exp_br_id = expadd.add_pkid ";
                    sql += " left join param expstate on expadd.add_state_id = expstate.param_pkid ";
                    sql += " left join bl notify on a.hbl_pkid = notify.bl_pkid ";
                    sql += " left join param liner on a.hbl_carrier_id = liner.param_pkid ";
                    sql += " left join param cmdty on a.hbl_commodity_id = cmdty.param_pkid ";
                    sql += " order by branch,mbl_date, mblno";


                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{COMPCODE}", company_code);

                    Con_Oracle = new DBConnection();
                    Con_Oracle.BeginTransaction();

                    sql1 = " delete from profitreport where type = 'SEA-EXPORT-FWD' and jv_year = " + year_code;
                    Con_Oracle.ExecuteNonQuery(sql1);

                    Con_Oracle.ExecuteNonQuery(sql);
                    sql = " update profitreport set profit = nvl(sell,0) - nvl(buy,0) where type ='SEA-EXPORT-FWD'";
                    Con_Oracle.ExecuteNonQuery(sql);

                    sql = " update profitreport set roi = profit / buy *100 where buy > 0 and type ='SEA-EXPORT-FWD'";
                    Con_Oracle.ExecuteNonQuery(sql);

                    sql = " UPDATE profitreport t1 SET(teu) = (SELECT t2.hbl_book_cntr_teu FROM hblm t2  WHERE t2.HBL_TYPE = 'MBL-SE' and t1.mblid = t2.hbl_pkid )  ";
                    sql += " WHERE t1.type = 'SEA-EXPORT-FWD' and t1.job_type = 'FCL' ";
                    Con_Oracle.ExecuteNonQuery(sql);

                    sql = " UPDATE profitreport t1 ";
                    sql += " SET(teu) = (SELECT t2.teu ";
                    sql += " FROM (select mblid, max(sino) as sino, max(teu) as teu ";
                    sql += " from profitreport where type = 'SEA-EXPORT-FWD' and job_type in('BUYERS CONSOLE', 'CONSOLE','FCL')";
                    sql += " group by mblid) t2  ";
                    sql += " WHERE t1.mblid = t2.mblid and t1.sino = t2.sino)  ";
                    sql += " WHERE t1.type = 'SEA-EXPORT-FWD' and t1.job_type in('BUYERS CONSOLE', 'CONSOLE','FCL')  ";
                    Con_Oracle.ExecuteNonQuery(sql);

                    // Narayanan Update - if mblno is blank then teu should be zero
                    sql = " update profitreport set teu = 0 where type ='SEA-EXPORT-FWD' and mblno is null";
                    Con_Oracle.ExecuteNonQuery(sql);

                    /*
                    sql = " update profitreport p1 set profit = ";
                    sql += " ( select   sum(nvl(p2.sell, 0) - nvl(p2.buy, 0)) * p1.hbl_chwt / sum(p2.hbl_chwt) ";
                    sql += " from profitreport p2 ";
                    sql += "  where p1.mblid = p2.mblid )";
                    Con_Oracle.ExecuteNonQuery(sql);

                    sql = " update profitreport p1 set roi = profit / buy *100 where buy > 0";
                    Con_Oracle.ExecuteNonQuery(sql);
                    */

                    Con_Oracle.CommitTransaction();
                    Con_Oracle.CloseConnection();

                }

                if (type == "AIR-IMPORT")
                {
                    sql += " insert into profitreport( ";
                    sql += " type,jv_year,mblid,reportdate,branch,mblno,mbl_date,mblslno,folder_no, ";
                    sql += " job_type,sino,si_date,hblno,hbl_date, ";
                    sql += " buy_date,sell_date,exporter_name,consignee_name, ";
                    sql += " agent_name,sman_name,nomination,mbl_frt_status, ";
                    sql += " hbl_frt_status,mbl_chwt,hbl_chwt,mbl_grwt,hbl_grwt ";

                    sql += " ,pol,pod,pofd,orgin_country,discription,fin_year,inv_nos ";
                    sql += " ,shpr_location,shpr_state,shpr_created,notify,liner ";
                    sql += " ,ex_01,ex_02,ex_03,ex_04,ex_05 ";
                    sql += " ,in_01,in_02,in_03,in_04,in_05 ";
                    sql += " ,cost_dr,cost_cr,rebate_dr ";
                    sql += " ,buy,sell ";

                    sql += " ) ";


                    sql += " select 'AIR-IMPORT'," + year_code + ", mbl.hbl_pkid,";
                    sql += " to_char(hbl.rec_created_date,'DD-MON-YYYY') ,";
                    sql += " a.rec_branch_code as branch, ";
                    sql += " mbl.hbl_bl_no as mawb_no, ";
                    sql += " mbl.hbl_date as mawb_date,  ";
                    sql += " mbl.hbl_no as mblslno, ";
                    sql += " max(mbl.hbl_folder_no) as mbl_folder_no, ";
                    sql += " max(mbl.hbl_jobtype) as job_type,  ";
                    sql += " hbl.hbl_no as SINO, ";
                    sql += " hbl.rec_created_date as si_date, ";
                    sql += " hbl.hbl_bl_no as hawb_no, ";
                    sql += " hbl.hbl_date as hawb_date, ";
                    sql += " max(case when jvh_type = 'PN' then jvh_date else null end) as buy_date, ";
                    sql += " max(case when jvh_type = 'IN' then jvh_date else null end) as sell_date, ";
                    sql += " max(exp.cust_name) as exporter_name, ";
                    sql += " max(imp.cust_name) as consignee_name, ";
                    sql += " max(agnt.cust_name) as agent, ";
                    sql += " max(nvl(sman2.param_name, sman.param_name)) as sman_name, ";
                    sql += " max(exp.cust_nomination) as nomination, ";
                    sql += " max(mbl.hbl_terms) as mbl_frt_status, ";
                    sql += " max(hbl.hbl_terms) as hbl_frt_status, ";
                    sql += " max(mbl.hbl_chwt) as mbl_chwt, ";
                    sql += " max(hbl.hbl_chwt) as hbl_chwt, ";
                    sql += " max(mbl.hbl_grwt) as mbl_grwt, ";
                    sql += " max(hbl.hbl_grwt) as hbl_grwt, ";
                    sql += " max(pol.param_name) as pol,  ";
                    sql += " max(pod.param_name) as pod,  ";
                    sql += " max(pofd.param_name) as pofd, ";
                    sql += " max(cntry.param_name) as orgin_country, ";
                    sql += " hbl.hbl_remarks as discription, ";
                    sql += " a.jvh_year as fin_year, ";
                    sql += " max(hbl.hbl_ar_invnos) as inv_nos, ";
                    sql += " max(impaddr.add_city) as imp_city, ";
                    sql += " max(impstate.param_name) as imp_state, ";
                    sql += " max(imp.rec_created_date) as imp_created, ";
                    sql += " max(notify.bl_notify_name) as notify, ";
                    sql += " carr.param_name as liner, ";

                    sql += "  sum( case when jv_drcr = 'DR' and acc_main_code  in('1401') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	ex_1401,";
                    sql += "  sum( case when jv_drcr = 'DR' and acc_main_code  in('1402') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	ex_1402,";
                    sql += "  sum( case when jv_drcr = 'DR' and acc_main_code  in('1403') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	ex_1403,";
                    sql += "  sum( case when jv_drcr = 'DR' and acc_main_code  in('1404') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	ex_1404,";
                    sql += "  sum( case when jv_drcr = 'DR' and acc_main_code  in('1405') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	ex_1405,";

                    sql += "  sum( case when jv_drcr = 'CR' and acc_main_code  in('1401') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	in_1401,";
                    sql += "  sum( case when jv_drcr = 'CR' and acc_main_code  in('1402') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	in_1402,";
                    sql += "  sum( case when jv_drcr = 'CR' and acc_main_code  in('1403') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	in_1403,";
                    sql += "  sum( case when jv_drcr = 'CR' and acc_main_code  in('1404') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	in_1404,";
                    sql += "  sum( case when jv_drcr = 'CR' and acc_main_code  in('1405') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	in_1405,";

                    sql += "  sum( case when jvh_type  in('HO','IN-ES')  and jv_drcr = 'DR' and acc_code  in('1405001')  then ABS(ct_amount) else 0 end ) as  cost_dr,";
                    sql += "  sum( case when jvh_type  in('HO','IN-ES')  and jv_drcr = 'CR' and acc_code  in('1405001')  then ABS(ct_amount) else 0 end ) as 	cost_cr,";

                    sql += "  sum( case when jv_drcr = 'DR' and acc_code  in('1401100','1402100','1403100','1404100','1405100')  then ABS(ct_amount) else 0 end ) as 	rebate_dr,";

                    sql += "  sum( case when jv_drcr = 'DR' and acc_main_code in('1401','1402','1403','1404','1405') then ABS(ct_amount) else 0 end ) as buy,";
                    sql += "  sum( case when jv_drcr = 'CR' and acc_main_code in('1401','1402','1403','1404','1405')    then ABS(ct_amount) else 0 end ) as sell ";

                    sql += "  from ledgerh a";
                    sql += "  inner join ledgert b on  jvh_pkid = jv_parent_id ";
                    sql += "  inner join costcentert c on b.jv_pkid = c.ct_jv_id";
                    sql += "  inner join hblm hbl on c.ct_cost_id = hbl.hbl_pkid";
                    sql += "  left join hblm mbl on hbl.hbl_mbl_id = mbl.hbl_pkid";
                    sql += "  inner join acctm e on jv_acc_id = acc_pkid";
                    sql += "  left join customerm exp on hbl.hbl_exp_id = exp.cust_pkid";

                    sql += "  left join customerm imp on hbl.hbl_imp_id = imp.cust_pkid";
                    sql += "  left join custdet  cd on hbl.rec_branch_code = cd.det_branch_code and hbl.hbl_imp_id = cd.det_cust_id ";
                    sql += "  left join param sman on imp.cust_sman_id = sman.param_pkid";
                    sql += "  left join param sman2 on cd.det_sman_id = sman2.param_pkid";

                    sql += "  left join addressm impaddr on hbl.hbl_imp_br_id=impaddr.add_pkid";
                    sql += "  left join param impstate on impaddr.add_state_id=impstate.param_pkid";
                    sql += "  left join param carr on (hbl.hbl_carrier_id= carr.param_pkid)";

                    sql += "  left join customerm agnt on hbl.hbl_agent_id=agnt.cust_pkid";

                    sql += " left join param pol on mbl.hbl_pol_id = pol.param_pkid";
                    sql += " left join param pod on mbl.hbl_pod_id = pod.param_pkid";
                    sql += " left join param pofd on mbl.hbl_pofd_id = pofd.param_pkid";

                    sql += " left join bl notify on hbl.hbl_pkid = notify.bl_pkid";
                    sql += " left join param cntry on hbl.hbl_origin_country_id = cntry.param_pkid";

                    sql += "  where a.jvh_year = " + year_code  +" and a.rec_company_code = '{COMPCODE}' ";
                    //sql += " and a.rec_branch_code = '{BRCODE}' ";
                    sql += " and hbl.hbl_type = 'HBL-AI' ";
                    sql += "  group by a.rec_branch_code,mbl.hbl_pkid, mbl.hbl_no, mbl.hbl_bl_no, mbl.hbl_date, hbl.hbl_pkid, hbl.hbl_no, hbl.hbl_bl_no, hbl.rec_created_date,hbl.hbl_date,";
                    sql += "  hbl.hbl_remarks,carr.param_name,a.jvh_year ";


                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{COMPCODE}", company_code);


                    Con_Oracle = new DBConnection();


                    Con_Oracle.BeginTransaction();

                    sql1 = " delete from profitreport where type = 'AIR-IMPORT' and jv_year = " +  year_code;
                    Con_Oracle.ExecuteNonQuery(sql1);

                    Con_Oracle.ExecuteNonQuery(sql);

                    sql = " update profitreport set profit = nvl(sell,0) - nvl(buy,0) where type ='AIR-IMPORT'";
                    Con_Oracle.ExecuteNonQuery(sql);
                    sql = " update profitreport set roi = profit / buy *100 where buy > 0 and type ='AIR-IMPORT'";
                    Con_Oracle.ExecuteNonQuery(sql);


                    sql = " update profitreport p1 set roi = profit / buy *100 where buy > 0";
                    Con_Oracle.ExecuteNonQuery(sql);

                    Con_Oracle.CommitTransaction();

                    Con_Oracle.CloseConnection();

                }

                if (type  == "SEA-IMPORT")
                {
                    sql += " insert into profitreport( ";
                    sql += " type,jv_year,mblid,reportdate,branch,mblno,mbl_date,mblslno ";
                    sql += " ,mbl_status,folder_no,job_type,sino ";
                    sql += " ,si_date,hblno,hbl_date,buy_date ";
                    sql += " ,sell_date,exporter_name,consignee_name ";
                    sql += " ,agent_name,sman_name,nomination ";
                    sql += " ,mbl_frt_status,hbl_frt_status,hbl_grwt,hbl_ntwt,hbl_cbm ";
                    sql += " ,pol,pod,pofd,orgin_country,nature,discription,cntr ";
                    sql += " ,teu,fin_year,inv_nos,shpr_location,shpr_state ";
                    sql += " ,shpr_created,notify,liner,ex_01,ex_02,ex_03,ex_04,ex_05,ex_06,ex_07 ";
                    sql += " ,in_01,in_02,in_03,in_04,in_05,in_06,in_07,cost_dr,cost_cr,rebate_dr,buy,sell ";
                    sql += " ) ";

                    sql += " select 'SEA-IMPORT'," + year_code + ", mbl.hbl_pkid,";
                    sql += " to_char(hbl.rec_created_date,'DD-MON-YYYY') ,";
                    sql += " a.rec_branch_code as branch, ";
                    sql += " mbl.hbl_bl_no as mbl_no,";
                    sql += " mbl.hbl_date as mbl_date,";
                    sql += " mbl.hbl_no as mblslno,";
                    sql += " max(status.param_name) as mbl_status, ";
                    sql += " max(mbl.hbl_folder_no) as mbl_folder_no,";
                    sql += " max(mbl.hbl_jobtype) as job_type,";
                    sql += " hbl.hbl_no as SINO, ";
                    sql += " hbl.rec_created_date as si_date,";
                    sql += " hbl.hbl_bl_no as hbl_no,";
                    sql += " hbl.hbl_date as hbl_date,";
                    sql += " max(case when jvh_type = 'PN' then jvh_date else null end) as buy_date,";
                    sql += " max(case when jvh_type = 'IN' then jvh_date else null end) as sell_date,";
                    sql += " max(exp.cust_name) as exporter_name,";
                    sql += " max(imp.cust_name) as consignee_name,";
                    sql += " max(agnt.cust_name) as agent,";
                    sql += " max(nvl(sman2.param_name, sman.param_name)) as sman_name,";
                    sql += " max(exp.cust_nomination) as nomination,";
                    sql += " max(mbl.hbl_terms) as mbl_frt_status,";
                    sql += " max(hbl.hbl_terms) as hbl_frt_status,";
                    sql += " max(hbl.hbl_grwt) as hbl_grwt,";
                    sql += " max(hbl.hbl_ntwt) as hbl_ntwt,";
                    sql += " max(hbl.hbl_cbm) as hbl_cbm,";
                    sql += " max(pol.param_name) as pol, ";
                    sql += " max(pod.param_name) as pod, ";
                    sql += " max(pofd.param_name) as pofd,";
                    sql += " max(cntry.param_name) as orgin_country,";
                    sql += " max(hbl.hbl_nature) as nature, ";
                    sql += " hbl.hbl_remarks as discription,";
                    sql += " max(hbl.hbl_book_cntr) as cntr, ";
                    sql += " max(hbl.hbl_book_cntr_teu) as teu,";
                    sql += " a.jvh_year as fin_year,";
                    sql += " max(hbl.hbl_ar_invnos) as inv_nos,";
                    sql += " max(impaddr.add_city) as imp_city,";
                    sql += " max(impstate.param_name) as imp_state,";
                    sql += " max(imp.rec_created_date) as imp_created,";
                    sql += " max(notify.bl_notify_name) as notify,";
                    sql += " carr.param_name as liner,";




                    sql += "   sum( case when jv_drcr = 'DR' and acc_main_code  in('1301') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	ex_1301,";
                    sql += "   sum( case when jv_drcr = 'DR' and acc_main_code  in('1302') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	ex_1302,";
                    sql += "   sum( case when jv_drcr = 'DR' and acc_main_code  in('1303') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	ex_1303,";
                    sql += "   sum( case when jv_drcr = 'DR' and acc_main_code  in('1304') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	ex_1304,";
                    sql += "   sum( case when jv_drcr = 'DR' and acc_main_code  in('1305') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	ex_1305,";
                    sql += "   sum( case when jv_drcr = 'DR' and acc_main_code  in('1306') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	ex_1306,";
                    sql += "   sum( case when jv_drcr = 'DR' and acc_main_code  in('1307') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	ex_1307,";

                    sql += "   sum( case when jv_drcr = 'CR' and acc_main_code  in('1301') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	in_1301,";
                    sql += "   sum( case when jv_drcr = 'CR' and acc_main_code  in('1302') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	in_1302,";
                    sql += "   sum( case when jv_drcr = 'CR' and acc_main_code  in('1303') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	in_1303,";
                    sql += "   sum( case when jv_drcr = 'CR' and acc_main_code  in('1304') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	in_1304,";
                    sql += "   sum( case when jv_drcr = 'CR' and acc_main_code  in('1305') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	in_1305,";
                    sql += "   sum( case when jv_drcr = 'CR' and acc_main_code  in('1306') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	in_1306,";
                    sql += "   sum( case when jv_drcr = 'CR' and acc_main_code  in('1307') and jvh_type  not in('HO','IN-ES')  then ABS(ct_amount) else 0 end ) as 	in_1307,";

                    sql += "   sum( case when jvh_type  in('HO','IN-ES')  and jv_drcr = 'DR' and acc_main_code  in('1305')  then ABS(ct_amount) else 0 end ) as  cost_dr,";
                    sql += "   sum( case when jvh_type  in('HO','IN-ES')  and jv_drcr = 'CR' and acc_main_code  in('1305')  then ABS(ct_amount) else 0 end ) as 	cost_cr,";

                    sql += "   sum( case when jv_drcr = 'DR' and acc_code  in('1301100','1302100','1303100','1304100','1305100','1306100','1307100')  then ABS(ct_amount) else 0 end ) as 	rebate_dr,";

                    sql += "   sum( case when jv_drcr = 'DR' and acc_main_code in('1301','1302','1303','1304','1305','1306','1307') then ABS(ct_amount) else 0 end ) as buy,";
                    sql += "   sum( case when jv_drcr = 'CR' and acc_main_code in('1301','1302','1303','1304','1305','1306','1307')    then ABS(ct_amount) else 0 end ) as sell";



                    sql += "   from ledgerh a";
                    sql += "   inner join ledgert b on  jvh_pkid = jv_parent_id ";
                    sql += "   inner join costcentert c on b.jv_pkid = c.ct_jv_id";
                    sql += "   inner join hblm hbl on c.ct_cost_id = hbl.hbl_pkid";
                    sql += "   left join hblm mbl on hbl.hbl_mbl_id = mbl.hbl_pkid";
                    sql += "   inner join acctm e on jv_acc_id = acc_pkid";
                    sql += "   left join customerm exp on hbl.hbl_exp_id = exp.cust_pkid";
                    sql += "   left join customerm imp on hbl.hbl_imp_id = imp.cust_pkid";
                    sql += "   left join custdet  cd on hbl.rec_branch_code = cd.det_branch_code and hbl.hbl_imp_id = cd.det_cust_id ";
                    sql += "   left join param sman on imp.cust_sman_id = sman.param_pkid";
                    sql += "   left join param sman2 on cd.det_sman_id = sman2.param_pkid";
                    sql += "   left join addressm impaddr on hbl.hbl_imp_br_id=impaddr.add_pkid";
                    sql += "   left join param impstate on impaddr.add_state_id=impstate.param_pkid";
                    sql += "   left join param carr on hbl.hbl_carrier_id= carr.param_pkid";
                    sql += "   left join customerm agnt on hbl.hbl_agent_id=agnt.cust_pkid";

                    sql += " left join param pol on mbl.hbl_pol_id = pol.param_pkid";
                    sql += " left join param pod on mbl.hbl_pod_id = pod.param_pkid";
                    sql += " left join param pofd on mbl.hbl_pofd_id = pofd.param_pkid";

                    sql += " left join bl notify on hbl.hbl_pkid = notify.bl_pkid";
                    sql += " left join param status on mbl.hbl_status_id = status.param_pkid";
                    sql += " left join param cntry on hbl.hbl_origin_country_id = cntry.param_pkid";

                    sql += "   where a.jvh_year = " + year_code + " and  a.rec_company_code = '{COMPCODE}' ";

                    //sql += "   and a.rec_branch_code = '{BRCODE}' ";

                    sql += "   and hbl.hbl_type = 'HBL-SI'  ";
                    sql += "   group by a.rec_branch_code,mbl.hbl_pkid, mbl.hbl_no, mbl.hbl_bl_no, mbl.hbl_date, hbl.hbl_pkid,hbl.rec_created_date, hbl.hbl_no,hbl.hbl_date,";
                    sql += "  hbl.hbl_bl_no,hbl.hbl_remarks,carr.param_name,a.jvh_year";



                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{COMPCODE}", company_code);


                    Con_Oracle = new DBConnection();

                    Con_Oracle.BeginTransaction();

                    sql1 = " delete from profitreport where type = 'SEA-IMPORT' and jv_year = " + year_code;
                    Con_Oracle.ExecuteNonQuery(sql1);

                    Con_Oracle.ExecuteNonQuery(sql);
                    sql = " update profitreport set profit = nvl(sell,0) - nvl(buy,0) where type ='SEA-IMPORT'";
                    Con_Oracle.ExecuteNonQuery(sql);
                    sql = " update profitreport set roi = profit / buy *100 where buy > 0 and type ='SEA-IMPORT'";
                    Con_Oracle.ExecuteNonQuery(sql);

                    Con_Oracle.CommitTransaction();

                    Con_Oracle.CloseConnection();


                }

                if (type  == "SEA-EXPORT-CLEARING")
                {
                    sql += " insert into profitreport( ";
                    sql += " type,jv_year,mblid,reportdate,branch,mblno,mbl_date,job_type,buy_date ";
                    sql += " ,sell_date,exporter_name,consignee_name,agent_name,sman_name ";
                    sql += " ,nomination,mbl_frt_status,mbl_chwt,mbl_grwt,hbl_ntwt,hbl_cbm ";
                    sql += " ,pol,pod,pofd,cntr,inv_nos,shpr_location,shpr_state ";
                    sql += " ,shpr_created,notify,commodity,pod_country,buyer_name ";
                    sql += " ,ex_01,ex_02,ex_03,ex_07 ";
                    sql += " ,in_01,in_02,in_03,in_07 ";
                    sql += " ,rebate_dr,buy,sell ";
                    sql += ")";

                    sql += "  select 'SEA-EXPORT-CLR'," + year_code + ", j.job_pkid ,";
                    sql += " j.job_date,";
                    sql += " a.rec_branch_code as branch,";
                    sql += " j.job_docno as job_no,";
                    sql += " j.job_date as job_date,";
                    sql += "  max(j.job_type) as job_type, ";
                    sql += "  max(case when jv_drcr = 'DR' then jvh_date else null end ) as buy_date,";
                    sql += "  max(case when jv_drcr = 'CR' then jvh_date else null end ) as sell_date,";

                    sql += "  max( exp.cust_name) as exporter_name,";
                    sql += "  max( imp.cust_name) as consignee_name,";
                    sql += "  max(agent.cust_name) as agent, ";
                    sql += "  max( nvl(sman1.param_name ,nvl(sman2.param_name,sman.param_name))) as sman_name,";
                    sql += "  max(nvl(job_nomination,imp.cust_nomination)) as nomination,";
                    sql += "  max( j.job_terms) as job_frt_status,";
                    sql += "  max(j.job_chwt) as job_chwt, ";
                    sql += "  max(j.job_grwt) as job_grwt,";
                    sql += "  max(j.job_ntwt) as job_ntwt,";
                    sql += "  max(j.job_cbm) as job_cbm,";
                    sql += " max(pol.param_name) as pol, ";
                    sql += " max(pod.param_name) as pod, ";
                    sql += " max(pofd.param_name) as pofd, ";
                    sql += "  max(j.job_cntr_type) as job_cntr_type, ";
                    sql += "  max(h.hbl_ar_invnos) as job_invoice_nos, ";

                    sql += "  max(expadd.add_city) as shpr_location, ";
                    sql += "  max(expstate.param_name) as shpr_state, ";
                    sql += "  max(exp.rec_created_date) as shpr_created, ";
                    sql += "  max(notify.bl_notify_name) as notify, ";
                    sql += "  max(cmdty.param_name) as commodity, ";
                    sql += " max(podcntry.param_name) as pod_country, ";
                    sql += " max(buyer.cust_name) as buyer_name, ";
                    

                    sql += "  sum( case when jv_drcr = 'DR' and acc_main_code  in('1101')   then ABS(ct_amount) else 0 end ) as 	ex_1101,";
                    sql += "  sum( case when jv_drcr = 'DR' and acc_main_code  in('1102')   then ABS(ct_amount) else 0 end ) as 	ex_1102,";
                    sql += "  sum( case when jv_drcr = 'DR' and acc_main_code  in('1103')   then ABS(ct_amount) else 0 end ) as 	ex_1103,";
                    //sql += "  sum( case when jv_drcr = 'DR' and h.hbl_mbl_id is null and acc_main_code  in('1107')   then ABS(ct_amount) else 0 end ) as 	ex_1107,";
                    sql += "  sum( case when jv_drcr = 'DR' and (h.hbl_mbl_id is null or j.job_type ='CLEARING') and acc_main_code  in('1107')   then ABS(ct_amount) else 0 end ) as 	ex_1107,";

                    sql += "  sum( case when jv_drcr = 'CR' and acc_main_code  in('1101')   then ABS(ct_amount) else 0 end ) as 	in_1101,";
                    sql += "  sum( case when jv_drcr = 'CR' and acc_main_code  in('1102')   then ABS(ct_amount) else 0 end ) as 	in_1102,";
                    sql += "  sum( case when jv_drcr = 'CR' and acc_main_code  in('1103')   then ABS(ct_amount) else 0 end ) as 	in_1103,";
                    //sql += "  sum( case when jv_drcr = 'CR' and h.hbl_mbl_id is null and acc_main_code  in('1107')   then ABS(ct_amount) else 0 end ) as 	in_1107,";
                    sql += "  sum( case when jv_drcr = 'CR'  and (h.hbl_mbl_id is null or j.job_type ='CLEARING') and acc_main_code  in('1107')   then ABS(ct_amount) else 0 end ) as 	in_1107,";

                    sql += "  sum( case when jv_drcr = 'DR' and acc_code  in('1101100','1102100','1103100')  then ABS(ct_amount) else 0 end ) as 	rebate_dr,";

                    sql += "  sum( case when jv_drcr = 'DR' and (acc_main_code in('1101','1102','1103') or ( acc_main_code  in('1107') and ( h.hbl_mbl_id is null or j.job_type ='CLEARING')))  then ABS(ct_amount) else 0 end ) as 	buy,";
                    sql += "  sum( case when jv_drcr = 'CR' and (acc_main_code in('1101','1102','1103') or ( acc_main_code  in('1107') and ( h.hbl_mbl_id is null or j.job_type ='CLEARING')))  then ABS(ct_amount) else 0 end ) as 	sell";



                    sql += "  from ledgerh a";
                    sql += "  inner join ledgert b on  jvh_pkid = jv_parent_id ";
                    sql += "  inner join costcentert c on b.jv_pkid = c.ct_jv_id";
                    sql += "  inner join jobm j on c.ct_cost_id = j.job_pkid";
                    sql += "  inner join acctm e on jv_acc_id = acc_pkid";
                    sql += "  left join customerm exp on j.job_exp_id = exp.cust_pkid";
                    sql += "  left join custdet  cd on j.rec_branch_code = cd.det_branch_code and j.job_exp_id = cd.det_cust_id ";
                    sql += "  left join param sman on exp.cust_sman_id = sman.param_pkid";
                    sql += "  left join param sman2 on cd.det_sman_id = sman2.param_pkid";
                    sql += "  left join param sman1 on j.job_salesman_id = sman1.param_pkid";

                    sql += "  left join customerm imp on j.job_imp_id = imp.cust_pkid";

                    sql += " left join param pol on j.job_pol_id = pol.param_pkid";
                    sql += " left join param pod on j.job_pod_id = pod.param_pkid";
                    sql += " left join param pofd on j.job_pofd_id = pofd.param_pkid";

                    sql += " left  join hblm h on j.jobs_hbl_id = h.hbl_pkid";
                    sql += " left join customerm agent on h.hbl_agent_id = agent.cust_pkid";
                    sql += " left join addressm expadd on j.job_exp_br_id = expadd.add_pkid";
                    sql += " left join param expstate on expadd.add_state_id = expstate.param_pkid";
                    sql += " left join bl notify on h.hbl_pkid = notify.bl_pkid";
                    sql += " left join param cmdty on j.job_commodity_id = cmdty.param_pkid";
                    sql += " left join param podcntry on h.hbl_pod_country_id = podcntry.param_pkid ";
                    sql += " left join customerm buyer on h.hbl_buyer_id = buyer.cust_pkid ";

                    sql += "  where  a.jvh_year = " + year_code + " and a.rec_company_code = '{COMPCODE}' ";
                    //sql += "  and a.rec_branch_code = '{BRCODE}'";

                    sql += "  and j.rec_category = 'SEA EXPORT' ";
                    sql += "  group by a.rec_branch_code,j.job_pkid,j.job_date,j.job_docno ";



                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{COMPCODE}", company_code);


                    Con_Oracle = new DBConnection();
                    Con_Oracle.BeginTransaction();

                    sql1 = " delete from profitreport where type = 'SEA-EXPORT-CLR' and jv_year = " + year_code;
                    Con_Oracle.ExecuteNonQuery(sql1);

                    Con_Oracle.ExecuteNonQuery(sql);

                    sql = " update profitreport set sell = in_01 + in_02 + in_03 + in_07, buy = ex_01 + ex_02 + ex_03 + ex_07  where type ='SEA-EXPORT-CLR'";
                    Con_Oracle.ExecuteNonQuery(sql);


                    sql = " update profitreport set profit = nvl(sell,0) - nvl(buy,0) where type ='SEA-EXPORT-CLR'";
                    Con_Oracle.ExecuteNonQuery(sql);

                    sql = " update profitreport set roi = profit / buy *100 where buy > 0 and type ='SEA-EXPORT-CLR'";
                    Con_Oracle.ExecuteNonQuery(sql);

                    Con_Oracle.CommitTransaction();

                    Con_Oracle.CloseConnection();


                }

                if (type == "AIR-EXPORT-CLEARING")
                {

                    sql += " insert into profitreport( ";
                    sql += " type,jv_year,mblid,reportdate,branch,mblno,mbl_date,job_type,buy_date ";
                    sql += " ,sell_date,exporter_name,consignee_name,agent_name,sman_name ";
                    sql += " ,nomination,mbl_frt_status,mbl_chwt,mbl_grwt,hbl_ntwt,hbl_cbm ";
                    sql += " ,pol,pod,pofd,inv_nos,shpr_location,shpr_state ";
                    sql += " ,shpr_created,notify,commodity,pod_country,buyer_name ";
                    sql += " ,ex_01,ex_02,ex_03 ";
                    sql += " ,in_01,in_02,in_03 ";
                    sql += " ,rebate_dr,buy,sell ";
                    sql += ")";


                    sql += " select 'AIR-EXPORT-CLR', " + year_code + ",j.job_pkid,";
                    sql += " j.job_date,";
                    sql += " a.rec_branch_code as branch,";
                    sql += " j.job_docno as job_no,";
                    sql += " j.job_date as job_date,";
                    sql += "  max(j.job_type) as job_type, ";
                    sql += "  max(case when jv_drcr = 'DR' then jvh_date else null end ) as buy_date,";
                    sql += "  max(case when jv_drcr = 'CR' then jvh_date else null end ) as sell_date,";
                    sql += "  max( exp.cust_name) as exporter_name,";
                    sql += "  max( imp.cust_name) as consignee_name,";
                    sql += "  max(agent.cust_name) as agent, ";
                    sql += "  max( nvl(sman1.param_name,nvl(sman2.param_name,sman.param_name))) as sman_name,";
                    sql += "  max( nvl(job_nomination,imp.cust_nomination)) as nomination,";
                    sql += "  max( j.job_terms) as job_frt_status,";
                    sql += "  max(j.job_chwt) as job_chwt, ";
                    sql += "  max(j.job_grwt) as job_grwt,";
                    sql += "  max(j.job_ntwt) as job_ntwt,";
                    sql += "  max(j.job_cbm) as job_cbm,";
                    sql += " max(pol.param_name) as pol, ";
                    sql += " max(pod.param_name) as pod, ";
                    sql += " max(pofd.param_name) as pofd, ";
                    sql += "  max(h.hbl_ar_invnos) as job_invoice_nos, ";
                    sql += "  max(expadd.add_city) as shpr_location, ";
                    sql += "  max(expstate.param_name) as shpr_state, ";
                    sql += "  max(exp.rec_created_date) as shpr_created, ";
                    sql += "  max(notify.bl_notify_name) as notify, ";
                    sql += "  max(cmdty.param_name) as commodity, ";
                    sql += " max(podcntry.param_name) as pod_country, ";
                    sql += " max(buyer.cust_name) as buyer_name, ";

                    sql += "  sum( case when jv_drcr = 'DR' and acc_main_code  in('1201')   then ABS(ct_amount) else 0 end ) as 	ex_1201,";
                    sql += "  sum( case when jv_drcr = 'DR' and acc_main_code  in('1202')   then ABS(ct_amount) else 0 end ) as 	ex_1202,";
                    sql += "  sum( case when jv_drcr = 'DR' and acc_main_code  in('1203')   then ABS(ct_amount) else 0 end ) as 	ex_1203,";
                    sql += "  sum( case when jv_drcr = 'CR' and acc_main_code  in('1201')   then ABS(ct_amount) else 0 end ) as 	in_1201,";
                    sql += "  sum( case when jv_drcr = 'CR' and acc_main_code  in('1202')   then ABS(ct_amount) else 0 end ) as 	in_1202,";
                    sql += "  sum( case when jv_drcr = 'CR' and acc_main_code  in('1203')   then ABS(ct_amount) else 0 end ) as 	in_1203,";

                    sql += "  sum( case when jv_drcr = 'DR' and acc_code  in('1201100','1202100','1203100')  then ABS(ct_amount) else 0 end ) as 	rebate_dr,";

                    sql += "  sum( case when jv_drcr = 'DR' and acc_main_code in('1201','1202','1203')   then ABS(ct_amount) else 0 end ) as 	buy,";
                    sql += "  sum( case when jv_drcr = 'CR' and acc_main_code in('1201','1202','1203')  then ABS(ct_amount) else 0 end ) as 	sell";


                    sql += "  from ledgerh a";
                    sql += "  inner join ledgert b on  jvh_pkid = jv_parent_id ";
                    sql += "  inner join costcentert c on b.jv_pkid = c.ct_jv_id";
                    sql += "  inner join jobm j on c.ct_cost_id = j.job_pkid";
                    sql += "  inner join acctm e on jv_acc_id = acc_pkid";
                    sql += "  left join customerm exp on j.job_exp_id = exp.cust_pkid";
                    sql += "  left join custdet  cd on j.rec_branch_code = cd.det_branch_code and j.job_exp_id = cd.det_cust_id ";
                    sql += "  left join param sman on exp.cust_sman_id = sman.param_pkid";
                    sql += "  left join param sman2 on cd.det_sman_id = sman2.param_pkid";
                    sql += "  left join param sman1 on j.job_salesman_id = sman1.param_pkid";

                    sql += "  left join customerm imp on j.job_imp_id = imp.cust_pkid";

                    sql += " left join param pol on j.job_pol_id = pol.param_pkid";
                    sql += " left join param pod on j.job_pod_id = pod.param_pkid";
                    sql += " left join param pofd on j.job_pofd_id = pofd.param_pkid";

                    sql += " left  join hblm h on j.jobs_hbl_id = h.hbl_pkid";
                    sql += " left join customerm agent on h.hbl_agent_id = agent.cust_pkid";
                    sql += " left join bl notify on h.hbl_pkid = notify.bl_pkid";
                    sql += " left join addressm expadd on j.job_exp_br_id = expadd.add_pkid";
                    sql += " left join param expstate on expadd.add_state_id = expstate.param_pkid ";
                    sql += " left join param cmdty on j.job_commodity_id = cmdty.param_pkid";
                    sql += " left join param podcntry on h.hbl_pod_country_id = podcntry.param_pkid ";
                    sql += " left join customerm buyer on h.hbl_buyer_id = buyer.cust_pkid ";

                    sql += "  where a.jvh_year = "  + year_code +" and a.rec_company_code = '{COMPCODE}' ";
                    //sql += "  and a.rec_branch_code = '{BRCODE}'";

                    sql += "  and j.rec_category = 'AIR EXPORT' ";
                    sql += "  group by a.rec_branch_code,j.job_pkid,j.job_date,j.job_docno";

                    

                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{COMPCODE}", company_code);


                    Con_Oracle = new DBConnection();

                    Con_Oracle.BeginTransaction();

                    sql1 = " delete from profitreport where type = 'AIR-EXPORT-CLR' and jv_year = "  + year_code;
                    Con_Oracle.ExecuteNonQuery(sql1);

                    Con_Oracle.ExecuteNonQuery(sql);
                    sql = " update profitreport set profit = nvl(sell,0) - nvl(buy,0) where type ='AIR-EXPORT-CLR'";
                    Con_Oracle.ExecuteNonQuery(sql);
                    sql = " update profitreport set roi = profit / buy *100 where buy > 0 and type ='AIR-EXPORT-CLR'";
                    Con_Oracle.ExecuteNonQuery(sql);

                    sql = "update profitreport set teu= 0 where type = 'SEA-EXPORT-FWD' and job_type ='LCL'";
                    Con_Oracle.ExecuteNonQuery(sql);

                    Con_Oracle.CommitTransaction();

                    Con_Oracle.CloseConnection();

                }



                if (type == "GENERAL-JOB")
                {

                    sql = "";
                    sql += " insert into profitreport( ";
                    sql += " type,jv_year, mblid, reportdate, mbl_date, mblslno, branch, buy_date, sell_date, exporter_name, sman_name, ";
                    sql += " pol, pod, pofd, fin_year, inv_nos, shpr_location, shpr_state, shpr_created, buy, sell ";
                    sql += " ) ";
                    sql += " select 'GENERAL-JOB', " + year_code + ", mbl.hbl_pkid, mbl.hbl_date, mbl.hbl_date as mbl_date,mbl.hbl_no as mblslno, a.rec_branch_code as branch, ";
                    sql += " max(case when jvh_type = 'PN' then jvh_date else null end) as buy_date, ";
                    sql += " max(case when jvh_type = 'IN' then jvh_date else null end) as sell_date, ";
                    sql += " max(exp.cust_name) as exporter_name, ";
                    sql += " max(nvl(sman2.param_name, sman.param_name)) as sman_name, ";
                    sql += " max(pol.param_name) as pol,  ";
                    sql += " max(pod.param_name) as pod,  ";
                    sql += " max(pofd.param_name) as pofd,  ";
                    sql += " max(a.jvh_year) as fin_year,  ";
                    sql += " max(mbl.hbl_ar_invnos) as inv_nos,  ";
                    sql += " max(expadd.add_city) as shpr_location, ";
                    sql += " max(expstate.param_name) as shpr_state,  ";
                    sql += " max(exp.rec_created_date) as exp_created, ";

                    sql += " sum( case when jv_drcr = 'DR' then ABS(ct_amount) else 0 end) as buy, ";
                    sql += " sum( case when jv_drcr = 'CR' then ABS(ct_amount) else 0 end) as sell ";

                    sql += " from ledgerh a inner ";
                    sql += " join ledgert b on jvh_pkid = jv_parent_id" ;
                    sql += " inner join costcentert c on b.jv_pkid = c.ct_jv_id ";
                    sql += " inner join hblm mbl on c.ct_cost_id = mbl.hbl_pkid ";
                    sql += " inner join acctm e on jv_acc_id = acc_pkid ";
                    sql += " left join customerm exp on mbl.hbl_exp_id = exp.cust_pkid ";
                    sql += " left join custdet cd on mbl.rec_branch_code = cd.det_branch_code and mbl.hbl_exp_id = cd.det_cust_id ";
                    sql += " left join param sman on exp.cust_sman_id = sman.param_pkid ";
                    sql += " left join param sman2 on cd.det_sman_id = sman2.param_pkid ";
                    sql += " left join param pol on mbl.hbl_pol_id = pol.param_pkid ";
                    sql += " left join param pod on mbl.hbl_pod_id = pod.param_pkid ";
                    sql += " left join param pofd on mbl.hbl_pofd_id = pofd.param_pkid ";
                    sql += " left join addressm expadd on mbl.hbl_exp_br_id = expadd.add_pkid ";
                    sql += " left join param expstate on expadd.add_state_id = expstate.param_pkid ";
                    

                    sql += "  where a.jvh_year = " + year_code +  " and  a.rec_company_code = '{COMPCODE}' ";
                    //sql += "  and a.rec_branch_code = '{BRCODE}'";
                    sql += "  and hbl_type  = 'JOB-GN' ";
                    sql += "  group by a.rec_branch_code,  mbl.hbl_pkid, mbl.hbl_date, mbl.hbl_date ,mbl.hbl_no ";


                    sql = sql.Replace("{BRCODE}", branch_code);
                    sql = sql.Replace("{COMPCODE}", company_code);


                    Con_Oracle = new DBConnection();

                    Con_Oracle.BeginTransaction();

                    sql1 = " delete from profitreport where type = 'GENERAL-JOB' and jv_year =" + year_code;
                    Con_Oracle.ExecuteNonQuery(sql1);

                    Con_Oracle.ExecuteNonQuery(sql);
                    sql = " update profitreport set profit = nvl(sell,0) - nvl(buy,0) where type ='GENERAL-JOB'";
                    Con_Oracle.ExecuteNonQuery(sql);
                    sql = " update profitreport set roi = profit / buy *100 where buy > 0 and type ='GENERAL-JOB'";
                    Con_Oracle.ExecuteNonQuery(sql);

                    Con_Oracle.CommitTransaction();

                    Con_Oracle.CloseConnection();

                }





            }
            catch (Exception Ex)
            {
                if (Con_Oracle != null)
                    Con_Oracle.CloseConnection();
                throw Ex;
            }

            RetData.Add("type", "");
            return RetData;
        }

    }
}
