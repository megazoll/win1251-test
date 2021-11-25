#include "metastd.h"
#include "metautil.h"
#include "ioobj.h"
#include "metadata.h"
#include "metactrl.h"
#include "fileutil.h"
#include "metawnd.h"
#include "tblform.h"
#include "objdata.h"
#include "svdkadat.h"
#include "datavoc.h"
#include "params.h"

#include "access.h"

#include "stdmenid.h"
#include "tblid.h"
#include "datafld.h"
#include "metamsg.h"


#include "srch.h"
#include "kadr.h"
#include "viewutil.h"
#include "viewfax.h"
#include "cliutil.h"
#include "bdsave.h"
#include "tblfldid.h"

#include "prjdocsrch.h"
#include "dogovor_rekr.h"
#include "metamath.h"
#include "akt_rekr.h"
#include "voicerecr.h"

#include "intl\intl.h"
#include "pointer.h"
#include "ctrl\ExportToExcel.h"
#include "ctrl\MSOControl.h"
#include "kadrysalary.h"
#include "dogovor.h"
#include "newclass.h"

static const char* _GetVoiceRecrBasicName(sdword voicetype)
{
	return (voicetype==VOICERECRTYPE_PAY)?LS(IDS_06726):
		(voicetype==VOICERECRTYPE_PAYAFTERPREPAY)?LS(IDS_06724):
		(voicetype==VOICERECRTYPE_PREPAY)?LS(IDS_06728):
		(voicetype==VOICERECRTYPE_MONEYBACK)?LS(IDS_02024):
		(voicetype==VOICERECRTYPE_ONLYVOICEFAKT)?LS(IDS_06893):
		(voicetype==VOICERECRTYPE_ADDPAY)?LS(IDS_06725):
		(voicetype==VOICERECRTYPE_MONEYBACK_NEW)?LS(IDS_15508):
		LS(IDS_06723);
}
const char* _GetVoiceRecrName(sdword voicetype,const char* voicenum,const char* voicefaktnum,
						sdword voicedate,sdword paydate,sdword zakkadryorgkey)
{
	char* buffer=GetKBMemory();
	const char* basicname=_GetVoiceRecrBasicName(voicetype);
	const char* numstr=(voicetype==VOICERECRTYPE_ONLYVOICEFAKT)||(voicetype==VOICERECRTYPE_MONEYBACK_NEW)?voicefaktnum:voicenum;
	strcpy(buffer,basicname);
	if (numstr && numstr[0]) sprintf(strtail(buffer)," № %s",numstr);
	if (voicedate>0) sprintf(strtail(buffer),LS(IDS_01203),_db2viewdate(long2str(voicedate)));
	if (paydate>0) sprintf(strtail(buffer),LS(IDS_01201),_db2viewdate(long2str(paydate)));
	if (zakkadryorgkey>0) sprintf(strtail(buffer)," (%s)",_GetKadryOrgName(zakkadryorgkey));
	return buffer;
}
const char* _GetVoiceRecrName(const VoiceRecrData* data,int whatadd)
{
	return _GetVoiceRecrName(data->voicetype,data->voicenum,data->voicefaktnum,
								data->voicedate,data->paydate,
								(whatadd & 1)?data->aktdata->dogdata->zakazkey:0);
}

const char* _GetVoiceRecrStatStr(sdword voicestat,sdword statdate, int /*voiceType*/, sdword addpar, int akttype)
{
	char* buffer=GetKBMemory();
	buffer[0]=0;
	switch (voicestat)
	{
/*
		просто так написал какую то херню
*/
		case VOICERECR_STAT_SENDTOBUH: strcpy(buffer,LS(IDS_05143)); break;
		case VOICERECR_STAT_SETVOICEDATE: strcpy(buffer,LS(IDS_02430)); break;
		case VOICERECR_STAT_SENDTOZAK: (akttype == AKTRECRTYPE_LEASING) ? strcpy(buffer, LS(IDS_20249)) : strcpy(buffer,LS(IDS_05026)); break;
		case VOICERECR_STAT_ZAKRECV: strcpy(buffer,LS(IDS_05503)); break;
		case VOICERECR_STAT_PAYED: (akttype == AKTRECRTYPE_LEASING) ? strcpy(buffer, LS(IDS_20250)) : strcpy(buffer, LS(IDS_04806)); break;
		case VOICERECR_STAT_CLOSE: strcpy(buffer,LS(IDS_06036)); break;
		case VOICERECR_STAT_ANNUL: strcpy(buffer,LS(IDS_01683)); break;
		case VOICERECR_STAT_NEEDANNUL_BUH: strcpy(buffer,LS(IDS_03315)); break;
		case VOICERECR_STAT_ANNUL_BYBUH: strcpy(buffer,LS(IDS_01684)); break;
		case VOICERECR_STAT_SENDTOBUH_NEEDCHANGE: strcpy(buffer,LS(IDS_05144)); break;
		case VOICERECR_STAT_NEEDCOMMENT_FOR_BUH: strcpy(buffer,LS(IDS_03317)); break;
		case VOICERECR_MONEYBACKSTAT_NEW_WAIT_OSW: strcpy(buffer,LS(IDS_11577)); break;
		case VOICERECR_MONEYBACKSTAT_NEW_WAITSENDTOFRC: 
		case VOICERECR_MONEYBACKSTAT_WAITSENDTOFRC: 
			strcpy(buffer,LS(IDS_04754)); 
			break;
		case VOICERECR_MONEYBACKSTAT_NEW_SENDEDTOFRC:
		case VOICERECR_MONEYBACKSTAT_SENDEDTOFRC:
			strcpy(buffer,LS(IDS_05146)); 
			break;
		case VOICERECR_STAT_INWORK_BUH:
			strcpy(buffer,LS(IDS_16607)); 
addaddfio:
			if (addpar>0) sprintf(strtail(buffer)," %s",_GetFIO(addpar));
			break;
		case VOICERECR_STAT_INWORK_BUH_NEEDCHANGE:
			strcpy(buffer,LS(IDS_16608)); 
			goto addaddfio;
	}
	if (buffer[0]&&(statdate>0))
	{
		sprintf(strtail(buffer)," (%s)",_db2viewdate(long2str(statdate)));
	}
	return buffer;
}
const char* _GetVoiceRecrStatStrFromStatQuary(int selqryline, SelectQuary* pSelectQuary, int akttype)
{
	int voicestat=pSelectQuary->GetIntField(selqryline,VOICERECRSTAT_VOICESTAT);
	sdword statdate=pSelectQuary->GetIntField(selqryline,VOICERECRSTAT_STATDATE);
	sdword addpar=0;
	switch (voicestat)
	{
	case VOICERECR_STAT_INWORK_BUH: addpar=pSelectQuary->GetIntField(selqryline,VOICERECRSTAT_1_BUH_INWORK_BUH); break;
	case VOICERECR_STAT_INWORK_BUH_NEEDCHANGE: addpar=pSelectQuary->GetIntField(selqryline,VOICERECRSTAT_1_BUH_INWORK_BUH_NEEDCHANGE); break;
	}

	return _GetVoiceRecrStatStr(voicestat, statdate, -1, addpar, akttype);
}
const char* _GetVoiceRecrStatStrFromAnotherObjQuary(int selqryline, SelectQuary* pSelectQuary, int akttype)
{
	int voicestat=pSelectQuary->GetIntField(selqryline,VOICERECR_ANOTHEROBJ_VOICESTAT);
	sdword statdate=pSelectQuary->GetIntField(selqryline,VOICERECR_ANOTHEROBJ_STATDATE);
	sdword addpar=0;

	return _GetVoiceRecrStatStr(voicestat, statdate, -1, addpar, akttype);
}
sdword GetVoiceRecrValuta(const SelectQuary* pZayavkaList)
{
	sdword result = 0;
	const int numApps = GetSelectQuaryNumLine(pZayavkaList);
	for(int i = 0; i < numApps; ++i)
	{
		sdword valuta = pZayavkaList->GetIntField(i, DOGOVORRECR_ANOTHEROBJ_GONORAR_VALUTA);
		if (!result)
		{
			result = valuta;
			continue;
		}
		if (result != valuta)
			return 0;
	}
	return result;
}
const char* GetCommentForBuh(int voiceType,int voiceStat,int aktStat)
{
	const char* p="";
	switch (voiceStat)
	{
		case VOICERECR_STAT_SENDTOBUH: 
		case VOICERECR_STAT_INWORK_BUH: 
			switch (voiceType)
			{
				case VOICERECRTYPE_PAY:
					p=LS(IDS_02424); 
					break;
				case VOICERECRTYPE_PAYAFTERPREPAY:
					p=LS(IDS_02425); 
					break;
				case VOICERECRTYPE_PREPAY:
					p=LS(IDS_02427); 
					break;
				case VOICERECRTYPE_MONEYBACK:
					p=LS(IDS_01989);
					break;
				case VOICERECRTYPE_ONLYVOICEFAKT:
					p=LS(IDS_02428); 
					break;
				case VOICERECRTYPE_ADDPAY:
					p=LS(IDS_02426); 
					break;
				case VOICERECRTYPE_MONEYBACK_NEW:
					p=LS(IDS_11578); 
					break;
			}
			break;
		case VOICERECR_STAT_SENDTOBUH_NEEDCHANGE:
		case VOICERECR_STAT_INWORK_BUH_NEEDCHANGE:
			switch (voiceType)
			{
				case VOICERECRTYPE_PAY:
					p=LS(IDS_03514); 
					break;
				case VOICERECRTYPE_PAYAFTERPREPAY:
					p=LS(IDS_03516); 
					break;
				case VOICERECRTYPE_PREPAY:
					p=LS(IDS_03517); 
					break;
				case VOICERECRTYPE_MONEYBACK:
					p=LS(IDS_01989);
					break;
				case VOICERECRTYPE_ONLYVOICEFAKT:
					p=LS(IDS_03518); 
					break;
				case VOICERECRTYPE_ADDPAY:
					p=LS(IDS_03515); 
					break;
				case VOICERECRTYPE_MONEYBACK_NEW:
					p=LS(IDS_11579); 
					break;
			}
			break;
		case VOICERECR_STAT_SETVOICEDATE:
		case VOICERECR_STAT_SENDTOZAK:
			p=LS(IDS_06734);
			break;
		case VOICERECR_STAT_ZAKRECV:
			p=(voiceType==VOICERECRTYPE_ONLYVOICEFAKT)?LS(IDS_06744):
				(voiceType==VOICERECRTYPE_MONEYBACK_NEW)?LS(IDS_11580):
				LS(IDS_01992);
			break;
		case VOICERECR_STAT_PAYED:
			p=(voiceType==VOICERECRTYPE_MONEYBACK)||(voiceType==VOICERECRTYPE_MONEYBACK_NEW)?LS(IDS_02494):LS(IDS_06738);
			break;
		case VOICERECR_STAT_CLOSE:
			p=LS(IDS_06036);
			break;
		case VOICERECR_STAT_ANNUL:
			p=LS(IDS_06737);
			break;
		case VOICERECR_STAT_NEEDANNUL_BUH:
			p=LS(IDS_01695);
			break;
		case VOICERECR_STAT_ANNUL_BYBUH:
			p=LS(IDS_06732);
			break;
		case VOICERECR_STAT_NEEDCOMMENT_FOR_BUH:
			p=LS(IDS_06733);
			break;
		case VOICERECR_MONEYBACKSTAT_WAITSENDTOFRC:
			if ((aktStat==AKTRECR_STAT_ZAKSIGN)||(aktStat==AKTRECR_STAT_CLOSE)) p=LS(IDS_15471);
			else p=LS(IDS_06735);
			break;
		case VOICERECR_MONEYBACKSTAT_SENDEDTOFRC:
			p=LS(IDS_15473);
			break;
		case VOICERECR_MONEYBACKSTAT_NEW_WAIT_OSW:
			p=LS(IDS_11581);
			break;
		case VOICERECR_MONEYBACKSTAT_NEW_WAITSENDTOFRC:
			p=LS(IDS_05152);
			break;
		case VOICERECR_MONEYBACKSTAT_NEW_SENDEDTOFRC:
			p=LS(IDS_06736);
			break;
	}
	return p;
}

int GetVoiceRecrConditionStr(char* buffer,uint buffsz,const VoiceRecrData* data)
{
	buffer[0]=0;
	if (!data) return 0;

	char tmpBuf[KB]; *tmpBuf = 0;
	#if DAS_USLUGA()
	if (GetCondClassifStr(tmpBuf, SIZEOFARRAY(tmpBuf), OBJTYPE_VOICERECR, &data->classifCodeList))
		sz_strcat(tmpBuf, "\n", SIZEOFARRAY(tmpBuf));
	#endif // DAS_USLUGA()
	sdword valuta = LicenseAccess_FreeWork() ? data->valuta : data->aktdata->dogdata->valuta;
	char* valutastr=GetDogovorValutaName(valuta);
#define ADDLINE(val,prompt) if (data->val>0) sprintf(strtail(buffer),"%s: %s %s\n",prompt,DigitalInt2Text(data->val),valutastr);
	if ((data->voicetype==VOICERECRTYPE_MONEYBACK)||(data->voicetype==VOICERECRTYPE_MONEYBACK_NEW))
	{
		ADDLINE(voicesum,LS(IDS_02025));
		ADDLINE(gonorarsum,LS(IDS_02027));
		sz_strcat(buffer, tmpBuf, buffsz);
	}
	else if (data->voicetype==VOICERECRTYPE_PREPAY)
	{
		ADDLINE(voicesum,LS(IDS_06703));
		sz_strcat(buffer, tmpBuf, buffsz);
		sz_strcat(buffer, LS(IDS_06624), buffsz);
		MakePayDayString(tmpBuf,MetaLang::RUSSIAN,data->aktdata->dogdata->prepaysrok_type,data->aktdata->dogdata->prepaysrok_days,PADEG_NOMINATIVE);
		sz_strcat(buffer, tmpBuf, buffsz);
	}
	else
	{
		ADDLINE(voicesum,LS(IDS_06699));
		if(data->aktdata->type == AKTRECRTYPE_LEASING)
		{
			FreeWorkAktAddInfo * addInfo = data->aktdata->GetFWAktAddData();
			if(addInfo && (addInfo->GetRealGonorar() > 0))
			{
				sprintf(strtail(buffer),"%s: %s %s\n",
					SetMyCompanyNameStrInStr(LS(IDS_02483)),DigitalInt2Text(addInfo->GetRealGonorar()),valutastr);
			}
			ADDLINE(gonorarsum, LS(IDS_10990));
		}
		else
		{
			ADDLINE(gonorarsum,SetMyCompanyNameStrInStr(LS(IDS_02483)));
		}

		sz_strcat(buffer, tmpBuf, buffsz);
		sz_strcat(buffer, LS(IDS_06624), buffsz);
		MakePayDayString(tmpBuf,MetaLang::RUSSIAN,data->aktdata->dogdata->paysrok_type,data->aktdata->dogdata->paysrok_days,PADEG_NOMINATIVE);
		sz_strcat(buffer, tmpBuf, buffsz);
	}
#undef ADDLINE
	DELETESTR(valutastr);
	return 1;
}
void VoiceRecrData::LoadFrom(const VoiceRecrData* fromdata)
{
	if (!fromdata) return;
	aktdata->LoadFrom(fromdata->aktdata);
	aktkey=fromdata->aktkey;
	prjkey=fromdata->prjkey;
	valuta=fromdata->valuta;
	ancoraccountkey=fromdata->ancoraccountkey;zakazaccountkey=fromdata->zakazaccountkey;
	voicetype=fromdata->voicetype;voicesum=fromdata->voicesum;
	gonorarsum=fromdata->gonorarsum;voicedate=fromdata->voicedate;
	paydate=fromdata->paydate;
	paysum=fromdata->paysum;
	strcpy(voicenum,fromdata->voicenum);
	strcpy(voicefaktnum,fromdata->voicefaktnum);
	strcpy(comment,fromdata->comment);
	classifCodeList.LoadFrom(&fromdata->classifCodeList);
	upClassifCodeList.LoadFrom(&fromdata->upClassifCodeList);
}
VoiceRecrData::VoiceRecrData(const VoiceRecrData* fromdata)
{
	aktdata=NEW AktRecrData();
	aktkey=prjkey=ancoraccountkey=zakazaccountkey=valuta=0;
	voicetype=voicesum=gonorarsum=voicedate=paydate=paysum=0;
	voicenum[0]=voicefaktnum[0]=comment[0]=0;
	LoadFrom(fromdata);
}
VoiceRecrData::~VoiceRecrData(void)
{
	DELETE(aktdata);
}
static sdword Get_1_Positive_From_3(sdword v1,sdword v2,sdword v3)
{
	if (v1>0) return v1;
	if (v2>0) return v2;
	return v3;
}
int VoiceRecrData::PutTo(ParamContainer* mainparam,const VoiceRecrData* olddata)
{
	if (!olddata)
	{
		mainparam->addbyseq(VOICERECR_AKTKEY,aktkey);
		mainparam->addbyseq(VOICERECR_DOGKEY,aktdata->dogkey);
		mainparam->addbyseq(VOICERECR_ANCORKEY,aktdata->dogdata->ancorkey);
		mainparam->addbyseq(VOICERECR_ZAKAZKEY,aktdata->dogdata->zakazkey);
		mainparam->addbyseq(VOICERECR_ANCORFILIAL,aktdata->dogdata->ancorfilial);
		mainparam->addbyseq(VOICERECR_VOICETYPE,voicetype);
		mainparam->addbyseq(VOICERECR_PRJKEY,prjkey);
	}
#define SETMAINFLD(fld,var) if ((!olddata)||(var!=olddata->var)) mainparam->addbyseq(fld,var)
	SETMAINFLD(VOICERECR_ANC_ACCOUNTKEY,ancoraccountkey);
	SETMAINFLD(VOICERECR_ZAK_ACCOUNTKEY,zakazaccountkey);
	SETMAINFLD(VOICERECR_VOICESUM,voicesum);
	SETMAINFLD(VOICERECR_GONORARSUM,gonorarsum);
	SETMAINFLD(VOICERECR_VOICEDATE,voicedate);
	SETMAINFLD(VOICERECR_PAYDATE,paydate);
	SETMAINFLD(VOICERECR_PAYSUM,paysum);
	SETMAINFLD(VOICERECR_VOICENUM,voicenum);
	SETMAINFLD(VOICERECR_VOICEFAKTNUM,voicefaktnum);
	SETMAINFLD(VOICERECR_COMMENT,comment);
#undef SETMAINFLD
	return mainparam->GetParamMask()!=0;
}
int VoiceRecrData::SetDogovorKey(HWND hWnd,sdword key)
{
	if (!aktdata->SetDogovorKey(hWnd,key)) return 0;

	#if DAS_USLUGA()
	{
		DPtr<KeyList> codeList = SetObjTypeForClassifCodeList(CLASSTYPE_DOGOVORUSLUGA, CLASSTYPE_VOICEUSLUGA, &aktdata->dogdata->classifCodeList);
		upClassifCodeList.LoadFrom(codeList);
		if (!classifCodeList.GetNumKey() && (upClassifCodeList.GetNumKey() == 1))
			classifCodeList.LoadFrom(&upClassifCodeList);
	}
	#endif // DAS_USLUGA()

	if (aktdata->dogdata->IsPrepay())
	{
		voicetype=VOICERECRTYPE_PREPAY;
		/* SKOREX if(GetDogovorType() == DOGOVORRECRTYPE_PLACEMENT_KONTRAKT)
			voicesum=aktdata->dogdata->gonorar_prepay_value;
		else*/
			voicesum=aktdata->dogdata->gonorar_numpos*aktdata->dogdata->gonorar_prepay_value;
		
		gonorarsum=0; 
		ancoraccountkey=Get_1_Positive_From_3(aktdata->dogdata->ancoraccountkey1,aktdata->dogdata->ancoraccountkey2,aktdata->dogdata->ancoraccountkey3);
		zakazaccountkey=Get_1_Positive_From_3(aktdata->dogdata->zakazaccountkey1,aktdata->dogdata->zakazaccountkey2,aktdata->dogdata->zakazaccountkey3);
		//voicedate=dbcurdate;
		paydate=paysum=0;
		voicenum[0]=voicefaktnum[0]=comment[0]=0;
	}
	return 1;
}
int VoiceRecrData::SetAktKey(HWND hWnd,sdword key)
{
	aktkey=key;
	aktdata->LoadFrom(hWnd,aktkey);

	#if DAS_USLUGA()
	{
		DebugTrap(); // проверь, что тут всё ок и убери эту строку
		DPtr<KeyList> codeList = SetObjTypeForClassifCodeList(CLASSTYPE_AKTUSLUGA, CLASSTYPE_VOICEUSLUGA, &aktdata->classifCodeList);
		upClassifCodeList.LoadFrom(codeList);
		if (!classifCodeList.GetNumKey() && (upClassifCodeList.GetNumKey() == 1))
			classifCodeList.LoadFrom(&upClassifCodeList);
	}
	#endif // DAS_USLUGA()

	switch (aktdata->type)
	{
		/* SKOREX case AKTRECRTYPE_SKX_SL_RESULT:
		case AKTRECRTYPE_SKX_PLMT_RESULT:*/
		case AKTRECRTYPE_RESULT:
		//case AKTRECRTYPE_ANNUL:
		case AKTRECRTYPE_ANNUL_CANCELFEE:
		case AKTRECRTYPE_CONSULT:
		case AKTRECRTYPE_RESULT_72:
		#if DAS_FREEWORK()
		case AKTRECRTYPE_LEASING:
		#endif //DAS_FREEWORK()
			voicetype=VOICERECRTYPE_PAY;
			break;
		// SKOREX case AKTRECRTYPE_SKX_PLMT_RESULT_PREPAY:
		case AKTRECRTYPE_RESULT_PREPAY:
			voicetype=VOICERECRTYPE_PAYAFTERPREPAY;
			break;
		// SKOREX case AKTRECRTYPE_SKX_PLMT_ANNUL_PREPAY:
		case AKTRECRTYPE_ANNUL_PREPAY:
			voicetype=VOICERECRTYPE_ONLYVOICEFAKT;
			break;
		case AKTRECRTYPE_MONEYBACK:
			//voicetype=VOICERECRTYPE_MONEYBACK;
			voicetype=VOICERECRTYPE_MONEYBACK_NEW;
			break;
		case AKTRECRTYPE_REPLACE:
			voicetype=VOICERECRTYPE_ADDPAY;
			break;
	}
	if (voicetype<=0) return 0;
	prjkey=aktdata->prjkey;
	ancoraccountkey=Get_1_Positive_From_3(aktdata->dogdata->ancoraccountkey1,aktdata->dogdata->ancoraccountkey2,aktdata->dogdata->ancoraccountkey3);
	zakazaccountkey=Get_1_Positive_From_3(aktdata->dogdata->zakazaccountkey1,aktdata->dogdata->zakazaccountkey2,aktdata->dogdata->zakazaccountkey3);
	//voicedate=dbcurdate;
	paydate=paysum=0;
	voicenum[0]=voicefaktnum[0]=comment[0]=0;
	if ((voicetype==VOICERECRTYPE_MONEYBACK)||(voicetype==VOICERECRTYPE_MONEYBACK_NEW))
	{
		voicesum=aktdata->pay_sum_back;
		gonorarsum=aktdata->service_fee_back;
		if (voicetype==VOICERECRTYPE_MONEYBACK_NEW)
		{
			voicedate=aktdata->aktdate;
		}
	}
	else
	{
		voicesum=aktdata->pay_sum;
		gonorarsum=aktdata->service_fee_sum;
	}
	return 1;
}
void VoiceRecrData::LoadFrom(int voiceqryline,SelectQuary* pVoiceQuary,
				int aktqryline,SelectQuary* pAktQuary,int aktaddqryline,SelectQuary* pAktAddQuary,int dogqryline,SelectQuary* pDogQuary,int dogcondqryline,SelectQuary* pDogCondQuary,
				int updogqryline,SelectQuary* pUpDogQuary,int updogcondqryline,SelectQuary* pUpDogCondQuary,
				SelectQuary* pResultPplQuary, SelectQuary* pObjSrcCodeQuary,SelectQuary* pZayavkaList)
{
	aktdata->LoadFrom(aktqryline,pAktQuary,aktaddqryline,pAktAddQuary,dogqryline,pDogQuary,dogcondqryline,pDogCondQuary,
				updogqryline,pUpDogQuary,updogcondqryline,pUpDogCondQuary,
				pResultPplQuary,0,0,0,0,0,0,pObjSrcCodeQuary);
#define VOICEQRYINT(fldid) pVoiceQuary->GetIntField_MinNull(voiceqryline,fldid)
#define VOICEQRYSTR(fldid) pVoiceQuary->GetField(voiceqryline,fldid)
	aktkey=VOICEQRYINT(VOICERECR_AKTKEY);
	prjkey=VOICEQRYINT(VOICERECR_PRJKEY);
	ancoraccountkey=VOICEQRYINT(VOICERECR_ANC_ACCOUNTKEY);
	zakazaccountkey=VOICEQRYINT(VOICERECR_ZAK_ACCOUNTKEY);
	voicetype=VOICEQRYINT(VOICERECR_VOICETYPE);
	voicesum=VOICEQRYINT(VOICERECR_VOICESUM);
	gonorarsum=VOICEQRYINT(VOICERECR_GONORARSUM);
	voicedate=VOICEQRYINT(VOICERECR_VOICEDATE);
	paydate=VOICEQRYINT(VOICERECR_PAYDATE);
	paysum=VOICEQRYINT(VOICERECR_PAYSUM);
	sz_strcpy(voicenum,VOICEQRYSTR(VOICERECR_VOICENUM),sizeof(voicenum));
	sz_strcpy(voicefaktnum,VOICEQRYSTR(VOICERECR_VOICEFAKTNUM),sizeof(voicefaktnum));
	sz_strcpy(comment,VOICEQRYSTR(VOICERECR_COMMENT),sizeof(comment));
	valuta = GetVoiceRecrValuta(pZayavkaList);

	#if DAS_USLUGA()
	{
		const sdword voicekey = VOICEQRYINT(VOICERECR_OWNKEY);
		const sdword dogkey = VOICEQRYINT(VOICERECR_DOGKEY);
		DPtr<KeyList> codeList = GetObjClassifCodeList(OBJTYPE_VOICERECR, voicekey, VOICERECROBJ_FLAGCLASSIF);
		classifCodeList.LoadFrom(codeList);
		if (aktkey > 0)
		{
			codeList = GetObjClassifCodeList(OBJTYPE_VOICERECR, voicekey, VOICERECROBJ_AKTFLAGCLASSIF, OBJTYPE_AKTRECR);
			codeList = SetObjTypeForClassifCodeList(CLASSTYPE_AKTUSLUGA, CLASSTYPE_VOICEUSLUGA, codeList);
			upClassifCodeList.LoadFrom(codeList);
		}
		else if (dogkey > 0)
		{
			codeList = GetObjClassifCodeList(OBJTYPE_VOICERECR, voicekey, VOICERECROBJ_DOGFLAGCLASSIF, OBJTYPE_DOGOVORRECR);
			codeList = SetObjTypeForClassifCodeList(CLASSTYPE_DOGOVORUSLUGA, CLASSTYPE_VOICEUSLUGA, codeList);
			upClassifCodeList.LoadFrom(codeList);
		}
		else
			DebugTrap(); // такого быть не должно. Счёт выходит либо из договора, либо из акта

		if (!classifCodeList.GetNumKey() && (upClassifCodeList.GetNumKey() == 1))
			classifCodeList.LoadFrom(&upClassifCodeList);
	}
	#endif // DAS_USLUGA()

#undef VOICEQRYSTR
#undef VOICEQRYINT

}
void VoiceRecrData::LoadFrom(HWND hWnd,sdword voicekey)
{
	HWAIT waitH=IsWindow(hWnd)?_WaitLoadObjData(hWnd,OBJTYPE_VOICERECR,voicekey,VOICERECROBJ_ALLMASK):0;
	const VoiceRecrData* d=GetVoiceRecrCardData(voicekey);
	if (d) LoadFrom(d);
	_UnlockWaitObjData(waitH);
}

int VoiceRecrData::GetDogovorType() const
{
	return aktdata ? aktdata->GetDogovorType() : 0;
}

int AddPrepayVoiceRecr(HWND hWnd,sdword dogkey,sdword prjkey)
{
	char* errStr=0;
	VoiceRecrData voicedata;
	sdword voicekey=-1;

	if (!voicedata.SetDogovorKey(hWnd,dogkey)) return 0;
	if (!voicedata.aktdata->dogdata->IsPrepay())
	{
		errStr=STRDUP(ReplaceDogovor2ProtocolStr(LS(IDS_01811),voicedata.GetDogovorType()));
	}
	else
	{
		int cont=1;
		HWAIT waitH=_WaitLoadObjData(hWnd,OBJTYPE_DOGOVORRECR,dogkey,OBJPART2MASK(DOGRECROBJ_VOICE)|OBJPART2MASK(DOGRECROBJ_AKT)|OBJPART2MASK(DOGRECROBJ_PRJ));
		SelectQuary* pVoiceQuary=GetObjSelectQuary(OBJTYPE_DOGOVORRECR,dogkey,DOGRECROBJ_VOICE);
		SelectQuary* pAktQuary=GetObjSelectQuary(OBJTYPE_DOGOVORRECR,dogkey,DOGRECROBJ_AKT);
		SelectQuary* pPrjQuary=GetObjSelectQuary(OBJTYPE_DOGOVORRECR,dogkey,DOGRECROBJ_PRJ);
		if ((!pVoiceQuary)||(!pAktQuary)||(!pPrjQuary)) cont=0;
		else
		{
			int l,voiceStat,aktStat,aktType,numline=pVoiceQuary->GetNumLine();
			for (l=0;l<numline;l++)
			{
				if (pVoiceQuary->GetIntField(l,VOICERECR_VOICETYPE)==VOICERECRTYPE_PREPAY)
				{
					voiceStat=pVoiceQuary->GetIntField(l,VOICERECR_ANOTHEROBJ_VOICESTAT);
					if ((voiceStat!=VOICERECR_STAT_ANNUL)&&
						(voiceStat!=VOICERECR_STAT_NEEDANNUL_BUH)&&
						(voiceStat!=VOICERECR_STAT_ANNUL_BYBUH))
					{
						errStr=STRDUP(ReplaceDogovor2ProtocolStr(LS(IDS_05300),voicedata.GetDogovorType()));
						break;
					}
				}
			}
			if (!errStr)
			{
				numline=pAktQuary->GetNumLine();
				for (l=0;l<numline;l++)
				{
					aktStat=pAktQuary->GetIntField(l,AKTRECR_ANOTHEROBJ_AKTSTAT);
					aktType=pAktQuary->GetIntField(l,AKTRECR_AKTTYPE);
					if ((aktType!=AKTRECRTYPE_CONSULT)&&(aktStat!=AKTRECR_STAT_ANNUL))
					{
						errStr=STRDUP(ReplaceDogovor2ProtocolStr(LS(IDS_02801),voicedata.GetDogovorType()));
						break;
					}
				}
			}
			if ((!errStr)&&(prjkey<=0))
			{
				numline=pPrjQuary->GetNumLine();
				if (numline<=0) errStr=STRDUP(ReplaceDogovor2ProtocolStr(LS(IDS_02802),voicedata.GetDogovorType()));
				else
				{
					KeyList* prjlist=MakeKeyListFromSelectQuary(pPrjQuary,DOGOVORRECRPRJ_PRJKEY);
					prjkey=ChooseObjFromListBox(hWnd,LS(IDS_07123),OBJTYPE_PRJ,prjlist,-1,0);
					DELETE(prjlist);
					if (prjkey<=0) cont=0;
				}
			}
		}
		_UnlockWaitObjData(waitH);
		if ((!errStr)&&cont)
		{
			voicedata.prjkey=prjkey;
			EditVoiceRecrDlg dlg(hWnd,&voicedata);
			if (dlg.display()==IDOK) errStr=AddVoiceRecrToBD(hWnd,voicekey,&voicedata);
		}
	}
	StdProceedError(hWnd,errStr);
	ViewObject(OBJTYPE_VOICERECR,voicekey);
	return voicekey>0;
}
int AddVoiceRecrToAkt(HWND hWnd,sdword aktkey,int rundlg)
{
	char* errStr=0;
	VoiceRecrData voicedata;
	sdword voicekey=-1;

	if (!voicedata.SetAktKey(hWnd,aktkey)) 
		return 0;

	int cont=1;
	HWAIT waitH=_WaitLoadObjData(hWnd,OBJTYPE_AKTRECR,aktkey,OBJPART2MASK(AKTRECROBJ_VOICE));
	SelectQuary* pVoiceQuary=GetObjSelectQuary(OBJTYPE_AKTRECR,aktkey,AKTRECROBJ_VOICE);
	if (!pVoiceQuary) cont=0;
	else
	{
		int l,voiceStat,numline=pVoiceQuary->GetNumLine();
		for (l=0;l<numline;l++)
		{
			voiceStat=pVoiceQuary->GetIntField(l,VOICERECR_ANOTHEROBJ_VOICESTAT);
			if ((voiceStat!=VOICERECR_STAT_ANNUL)&&
				(voiceStat!=VOICERECR_STAT_NEEDANNUL_BUH)&&
				(voiceStat!=VOICERECR_STAT_ANNUL_BYBUH))
			{
				break;
			}
		}
		if (l<numline)
		{
			if ((voicedata.voicetype==VOICERECRTYPE_ONLYVOICEFAKT)||(voicedata.voicetype==VOICERECRTYPE_MONEYBACK_NEW))
			{
				ErrorBox(hWnd,LS(IDS_05292));
				cont=0;
			}
			else
			{
				cont=IDYES==YesNoBox(hWnd,LS(IDS_05299));
			}
		}
	}
	_UnlockWaitObjData(waitH);
	if (!cont) return 0;

	int addfl=1;
	if (rundlg)
	{
		EditVoiceRecrDlg dlg(hWnd,&voicedata);
		addfl=(dlg.display()==IDOK);
	}
	if (addfl)
	{
		errStr=AddVoiceRecrToBD(hWnd,voicekey,&voicedata);
		StdProceedError(hWnd,errStr);
	}
	ViewObject(OBJTYPE_VOICERECR,voicekey);
	return voicekey>0;
}
void AddVoiceRecrAndNeedAnnulOld(HWND hWnd,sdword& newvoicekey,sdword oldvoicekey)
{
	if (!VerifyExistInnerDocInVoice(hWnd,oldvoicekey)) return;
	char* errStr=0;
	VoiceRecrData* voicedata=LoadVoiceRecrCardData(hWnd,oldvoicekey);
	if (!voicedata) return;
//	sdword aktkey,ancoraccountkey,zakazaccountkey;
	voicedata->voicedate=voicedata->paydate=voicedata->paysum=0;
	voicedata->voicenum[0]=voicedata->voicefaktnum[0]=voicedata->comment[0]=0;
	EditVoiceRecrDlg dlg(hWnd,voicedata);
	if (dlg.display()==IDOK)
	{
		errStr=AddVoiceRecrToBD(hWnd,newvoicekey,voicedata,oldvoicekey);
		if (StdProceedError(hWnd,errStr))
		{
			RunSimpleVoiceStatCommand(hWnd,IDM_VOICERECR_SETSTAT_NEEDANNUL_BUH,oldvoicekey,1);
		}
	}
	DELETE(voicedata);
}
int EditVoiceRecrCard(HWND hWnd,sdword voicekey,int dlgtype)
{
	VoiceRecrData* voicedata=LoadVoiceRecrCardData(hWnd,voicekey);
	if (!voicedata) return 0;
	VoiceRecrData olddata(voicedata);
	ParamContainer mainparam;
	int retval=0;
	EditVoiceRecrDlg dlg(hWnd,voicedata,dlgtype);
	if (dlg.display()==IDOK)
	{
		char* errStr=SaveVoiceRecrToBD(hWnd,voicekey,voicedata,&olddata);
		retval=StdProceedError(hWnd,errStr);
	}
	DELETE(voicedata);
	return retval;
}
void SetVoiceNum(HWND hWnd,sdword voicekey)
{
	VoiceRecrData* voicedata=LoadVoiceRecrCardData(hWnd,voicekey);
	if (!voicedata) return;
	VoiceRecrData olddata(voicedata);
	ParamContainer mainparam;
	char* errStr=0;
	char voicefile[256],voicefaktfile[256];
	int voicestat=GetVoiceRecrStat(hWnd,voicekey);
	voicefile[0]=voicefaktfile[0]=0;
	//if (voicestat==VOICERECR_STAT_SENDTOBUH) voicedata->voicedate=dbcurdate;
	int dlgtype=((voicedata->voicetype==VOICERECRTYPE_ONLYVOICEFAKT)||
				(voicedata->voicetype==VOICERECRTYPE_MONEYBACK_NEW))?3:1;
	EditVoiceRecrDlg dlg(hWnd,voicedata,dlgtype,voicefile,voicefaktfile);
	if (dlg.display()==IDOK)
	{
		errStr=SaveVoiceRecrToBD(hWnd,voicekey,voicedata,&olddata);
		if (errStr) goto done;
		if (voicefile[0])
		{
			errStr=AddBinaryDocToObjLikeDogovor(hWnd,OBJTYPE_VOICERECR,voicekey,voicefile,
						VOICERECR_DOCTYPE_VOICEEXCEL,GetFileExt(voicefile),LS(IDS_06799));
			if (errStr) goto done;
		}
		if (voicefaktfile[0])
		{
			errStr=AddBinaryDocToObjLikeDogovor(hWnd,OBJTYPE_VOICERECR,voicekey,voicefaktfile,
						VOICERECR_DOCTYPE_VOICEFAKTEXCEL,GetFileExt(voicefaktfile),LS(IDS_06800));
			if (errStr) goto done;
		}
		if ((voicestat==VOICERECR_STAT_SENDTOBUH)||(voicestat==VOICERECR_STAT_SENDTOBUH_NEEDCHANGE)||
			(voicestat==VOICERECR_STAT_INWORK_BUH)||(voicestat==VOICERECR_STAT_INWORK_BUH_NEEDCHANGE))
		{
			errStr=SetVoiceRecrStat(hWnd,voicekey,VOICERECR_STAT_SETVOICEDATE,voicedata,0);
			if (errStr) goto done;
		}

		// Нотификация ответственному и исполнителю о выставлении счета
		VoiceRecrAddData* addData = GetVoiceRecrAddData(voicekey);
		ListContainer* otvList = addData ? addData->GetOtvData() : 0;
		if(otvList)
		{
			KeyList sotrList;
			for(int num = 0; num < otvList->GetNumLine(); num++)
				sotrList.AddKey(otvList->GetKey(num));

			errStr=SendMsgToSotrList(hWnd, &sotrList, LS(IDS_16498), OBJTYPE_VOICERECR, voicekey); 
			if (errStr) goto done;
		}

		otvList = addData ? addData->GetSecrData() : 0;
		if(otvList)
		{
			KeyList sotrList;
			for(int num = 0; num < otvList->GetNumLine(); num++)
				sotrList.AddKey(otvList->GetKey(num));
			
			if(sotrList.GetNumKey() > 0)
				errStr=SendMsgToSotrList(hWnd, &sotrList, LS(IDS_16539), OBJTYPE_VOICERECR, voicekey); 
			if (errStr) goto done;
		}
	}
done:
	DELETE(voicedata);
	StdProceedError(hWnd,errStr);
}
extern char* SetVoiceRecrPayStat(HWND hWnd,int newstat,sdword voicekey);
void SetVoicePayed(HWND hWnd,sdword voicekey)
{
	VoiceRecrData* voicedata=LoadVoiceRecrCardData(hWnd,voicekey);
	if (!voicedata) return;
	VoiceRecrData olddata(voicedata);
	char* errStr=0;
	int voicestat=GetVoiceRecrStat(hWnd,voicekey);
	int newPayStat=-1;
	if ((voicestat==VOICERECR_STAT_ZAKRECV)||(voicestat==VOICERECR_MONEYBACKSTAT_SENDEDTOFRC)||
		(voicestat==VOICERECR_MONEYBACKSTAT_NEW_SENDEDTOFRC))
	{
		voicedata->paydate=0; //dbcurdate;
		voicedata->paysum=voicedata->voicesum;
	}
	EditVoiceRecrDlg dlg(hWnd,voicedata,2,0,0);
	if (dlg.display()==IDOK)
	{
		errStr=SaveVoiceRecrToBD(hWnd,voicekey,voicedata,&olddata);
		if (errStr) goto done;
		if ((voicestat==VOICERECR_STAT_ZAKRECV)||(voicestat==VOICERECR_MONEYBACKSTAT_SENDEDTOFRC)||
			(voicestat==VOICERECR_MONEYBACKSTAT_NEW_SENDEDTOFRC))
		{
			errStr=SetVoiceRecrStat(hWnd,voicekey,VOICERECR_STAT_PAYED,voicedata,0);
			if (errStr) goto done;
			if ((voicestat==VOICERECR_MONEYBACKSTAT_SENDEDTOFRC)||(voicestat==VOICERECR_MONEYBACKSTAT_NEW_SENDEDTOFRC))
			{
				newPayStat=VOICERECR_PAYSTAT_MONEYBACKOK;
			}
		}
		if ((newPayStat<0)&&
			((voicedata->voicetype==VOICERECRTYPE_PAY)||(voicedata->voicetype==VOICERECRTYPE_PAYAFTERPREPAY)||
					(voicedata->voicetype==VOICERECRTYPE_PREPAY)||(voicedata->voicetype==VOICERECRTYPE_ADDPAY)))
		{
			if (voicedata->voicesum==voicedata->paysum) newPayStat=VOICERECR_PAYSTAT_PAYOK;
			else if (voicedata->voicesum>voicedata->paysum) newPayStat=VOICERECR_PAYSTAT_PARTPAY;
			else if (voicedata->voicesum<voicedata->paysum) newPayStat=VOICERECR_PAYSTAT_OVERPAY;
		}
		if (newPayStat>0) errStr=SetVoiceRecrPayStat(hWnd,newPayStat,voicekey);
	}
done:
	DELETE(voicedata);
	StdProceedError(hWnd,errStr);
}
char* SetVoicePayedFrom1C(HWND hWnd,sdword voicekey,sdword paysum,sdword paydate)
{
	char* errStr=0;
	VoiceRecrData* voicedata=LoadVoiceRecrCardData(hWnd,voicekey);
	if (voicedata)
	{
		VoiceRecrData olddata(voicedata);
		int newPayStat=-1;
		voicedata->paydate=paydate;
		voicedata->paysum=paysum;
		errStr=SaveVoiceRecrToBD(hWnd,voicekey,voicedata,&olddata);
		if (errStr) goto done;
		errStr=SetVoiceRecrStat(hWnd,voicekey,VOICERECR_STAT_PAYED,voicedata,1);
		if (errStr) goto done;
		if (voicedata->voicesum==voicedata->paysum) newPayStat=VOICERECR_PAYSTAT_PAYOK;
		else if (voicedata->voicesum>voicedata->paysum) newPayStat=VOICERECR_PAYSTAT_PARTPAY;
		else if (voicedata->voicesum<voicedata->paysum) newPayStat=VOICERECR_PAYSTAT_OVERPAY;
		if (newPayStat>0) errStr=SetVoiceRecrPayStat(hWnd,newPayStat,voicekey);
	}
done:
	DELETE(voicedata);
	return errStr;
}
extern char* FillPplEnsiKeyFromMetaKey(HWND hWnd,sdword*& bosskey,KeyList* pplKeyList);
extern sdword GetCFOClassStrCode(sdword key,int classnum);
#if FW_GONORAR_NEW()
static void AdjustTblLineValueByDelta(TblData * pTbl, int line, int fld, sdword delta)
{
	if(line >= 0 && delta != 0)
	{
		char buf[16] = {0};
		double gonorar = Str2Double(pTbl->GetValue(line, fld));
		gonorar += DigitalInt2Double(delta);
		pTbl->SetValue(line,fld,Double2Str(buf,gonorar,DOUBLETYPE_MONEY));
	}
}
#endif //FW_GONORAR_NEW	
UnsortedKeyList* MakeUnsortedKeyListFromSelectQuaryWithObjKeyList(const SelectQuary* pSelectQuary,int fromfld,const KeyList* objkeyList,int objkeyfld)
{
	int numline=GetSelectQuaryNumLine(pSelectQuary);
	UnsortedKeyList * ulist = NEW UnsortedKeyList(numline);
	for (int l=0;l<numline;l++) 
	{
		if (objkeyList->IsKeyExist(pSelectQuary->GetIntField(l,objkeyfld))) ulist->AddKey(pSelectQuary->GetIntField(l,fromfld));
	}
	return ulist;
}
static DataBaseData* MakeVoiceRecrZayavkaWinWordData(HWND hWnd,VoiceRecrData* data,SelectQuary* pSotrQuary)
{
	DataBaseData* pData=MakeDialogData("VoiceRecrZayavkaData");
	SetDateDelimitor('.');

	DblKeyList ancororgfld;
	ancororgfld.AddDbl(VOICERECR_MAIN_ANCORNAME,      KADRYFILIALSTR_ORGFULLNAME);
	ancororgfld.AddDbl(VOICERECR_MAIN_ANCORKADRY_INN, KADRYFILIALSTR_INN        );
	ancororgfld.AddDbl(VOICERECR_MAIN_ANCORKADRY_KPP, KADRYFILIALSTR_KPP        );
	DblKeyList zakazorgfld;
	zakazorgfld.AddDbl(VOICERECR_MAIN_ZAKAZNAME,      KADRYFILIALSTR_ORGFULLNAME);
	zakazorgfld.AddDbl(VOICERECR_MAIN_ZAKAZKADRY_INN, KADRYFILIALSTR_INN        );
	zakazorgfld.AddDbl(VOICERECR_MAIN_ZAKAZKADRY_KPP, KADRYFILIALSTR_KPP        );
	FillOrgStrInTbl(hWnd,DOGOVORLANG_1(data->aktdata->dogdata->dwlang),data->aktdata->dogdata->ancorkey,
		data->aktdata->dogdata->ancorfilial,pData->GetSimpleData(TBL_VOICERECR_MAIN),ancororgfld);
	FillOrgStrInTbl(hWnd,DOGOVORLANG_1(data->aktdata->dogdata->dwlang),data->aktdata->dogdata->zakazkey,
		-1,pData->GetSimpleData(TBL_VOICERECR_MAIN),zakazorgfld);
	pData->SetIntValue(VOICERECR_MAIN_ANCORKADRYCODE,data->aktdata->dogdata->ancorkey);
	pData->SetIntValue(VOICERECR_MAIN_ZAKAZKADRYCODE,data->aktdata->dogdata->zakazkey);
	pData->SetValue(VOICERECR_MAIN_POSNAME,data->aktdata->posname);
	pData->SetValue(VOICERECR_MAIN_VOICE_TYPE,_GetVoiceRecrBasicName(data->voicetype));
	pData->SetValue(VOICERECR_MAIN_USLUGA_TYPE,combostr("SingleDogovorPosLevel",data->aktdata->dogdata->poslevel));
	double aktsum=0,voicesum=0,gonorar=0;
	if ((data->voicetype==VOICERECRTYPE_MONEYBACK)||(data->voicetype==VOICERECRTYPE_MONEYBACK_NEW))
	{
		aktsum=-DigitalInt2Double(data->aktdata->service_fee_back);
		voicesum=-DigitalInt2Double(data->voicesum);
		gonorar=-DigitalInt2Double(data->gonorarsum);
	}
	else
	{
		aktsum=DigitalInt2Double(data->aktdata->service_fee_sum);
		voicesum=DigitalInt2Double(data->voicesum);
		gonorar=DigitalInt2Double(data->gonorarsum);
	}
	char buffer[256],buf[16],buf1[16];
	char cfobuffer[16];
	sprintf(buffer,Const::Fmt::s_s,Double2Str(buf,aktsum,DOUBLETYPE_MONEY),_GetDogovorValutaName(data->aktdata->dogdata->valuta));
	pData->SetValue(VOICERECR_MAIN_AKT_SERVICEFEE,buffer);
	aktsum=aktsum*(1+DigitalInt2Double(GetNDSPercentDigitalInt(data->aktdata->aktdate,data->aktdata->dogdata->residentType))/100.0);
	sprintf(buffer,Const::Fmt::s_s,Double2Str(buf,aktsum,DOUBLETYPE_MONEY),_GetDogovorValutaName(data->aktdata->dogdata->valuta));
	pData->SetValue(VOICERECR_MAIN_AKT_SUM,buffer);
	pData->SetFloatValue(VOICERECR_MAIN_VOICE_SUM,voicesum,DOUBLETYPE_MONEY);
	pData->SetValue(VOICERECR_MAIN_COMMENT,data->comment);
	pData->SetFloatValue(VOICERECR_MAIN_VOICE_GONORAR_SUM, gonorar, DOUBLETYPE_MONEY);

	int n,l,numline=GetSelectQuaryNumLine(pSotrQuary);
	sdword cfokey,sotrkey;
	udword needMask,index;
	sdword percent;
	double gon,gonwithnds;
	KeyList *pSotrList=0;
	KeyList *pCFOList=0;
	KeyList *pOtvCFOList=0;
	sdword* sotrBossKey=0;
	if (numline>0)
	{
		pSotrList=MakeKeyListFromSelectQuary(pSotrQuary,LIKEDOGOVORRECRSOTR_SOTRKEY);
		pCFOList=MakeKeyListFromSelectQuary(pSotrQuary,LIKEDOGOVORRECRSOTR_OTVCFO);
		if (NUMKEYINKEYLIST(pCFOList)>0)
		{
			pOtvCFOList=CFOKeyList2ChiefCFOKeyList(pCFOList);
			if (NUMKEYINKEYLIST(pOtvCFOList)>0)
			{
				if (pSotrList) ORKeyList(pSotrList,pOtvCFOList);
				else
				{
					pSotrList=pOtvCFOList;
					pOtvCFOList=0;
				}
			}
		}
		if (NUMKEYINKEYLIST(pSotrList)>0)
		{
			char* errStr=FillPplEnsiKeyFromMetaKey(hWnd,sotrBossKey,pSotrList);
			DELETESTR(errStr);
		}

		ResortSelectQuaqryByObjName(LIKEDOGOVORRECRSOTR_SOTRKEY,OBJTYPE_PEOPLE,pSotrQuary);

		const bool bIsFreeWorkGenDog = (LicenseAccess_FreeWork() && (data->GetDogovorType() == DOGOVORRECRTYPE_LEASING_GEN));

		TblData* pTbl=pData->GetTblData(TBL_VOICERECR_SOTR);
		if (pTbl)
		{
			needMask=ENUM2BIT(DOGOVORRECR_SORTFUNC_OTV);
			for (l=0;l<numline;l++)
			{
				if (pSotrQuary->GetIntField_MinNull(l,LIKEDOGOVORRECRSOTR_SOTRFUNC) & needMask)
				{
					n=pTbl->GetNumLine();
					pTbl->InsertLine();
					sotrkey=pSotrQuary->GetIntField(l,LIKEDOGOVORRECRSOTR_SOTRKEY);
					cfokey=pSotrQuary->GetIntField(l,LIKEDOGOVORRECRSOTR_OTVCFO);
					pTbl->SetValue(n,VOICERECR_SOTR_FUNC,bIsFreeWorkGenDog?"Исполнитель":"Ответственный за счет");
					pTbl->SetValue(n,VOICERECR_SOTR_FIO,_GetFullFIO(sotrkey));
					if (cfokey>0)
					{
						sprintf(cfobuffer,"%09d",cfokey);
						pTbl->SetValue(n,VOICERECR_SOTR_CFOCODE,cfobuffer);
						//pTbl->SetIntValue(n,VOICERECR_SOTR_CFOCODE,cfokey);
					}
					if (sotrBossKey && pSotrList->IsKeyExist(sotrkey,&index))
					{
						sprintf(cfobuffer,"%09d",sotrBossKey[index]);
						pTbl->SetValue(n,VOICERECR_SOTR_BOSSCODE,cfobuffer);
						//pTbl->SetIntValue(n,VOICERECR_SOTR_BOSSCODE,sotrBossKey[index]);
					}
					if (cfokey>0)
					{
						pTbl->SetValue(n,VOICERECR_SOTR_CFO,_GetCFOName(cfokey));
						pData->SetValue(VOICERECR_MAIN_CFO_WHATDOCOMPANY,_GetCFOClassStr(cfokey,CFOVOCAB_CFOCLASS_WHATDOCOMPANY));
						sotrkey=CFOKey2ChiefCFOKey(cfokey);
						if (sotrkey>0)
						{
							pData->SetValue(VOICERECR_MAIN_CFO_CHIEF,_GetFullFIO(sotrkey));
							if (sotrBossKey && pSotrList->IsKeyExist(sotrkey,&index))
							{
								sprintf(cfobuffer,"%09d",sotrBossKey[index]);
								pData->SetValue(VOICERECR_MAIN_CFO_CHIEFCODE,cfobuffer);
							}
						}
						else
						{
							pData->SetValue(VOICERECR_MAIN_CFO_CHIEF,_GetCFOClassStr(cfokey,CFOVOCAB_CFOCLASS_CFOCHIEF));
						}
					}
				}
			}
			needMask=ENUM2BIT(DOGOVORRECR_SORTFUNC_SECR);
			for (l=0;l<numline;l++)
			{
				if (pSotrQuary->GetIntField_MinNull(l,LIKEDOGOVORRECRSOTR_SOTRFUNC) & needMask)
				{
					n=pTbl->GetNumLine();
					pTbl->InsertLine();
					pTbl->SetValue(n,VOICERECR_SOTR_FUNC,LS(IDS_03647));
					sotrkey=pSotrQuary->GetIntField(l,LIKEDOGOVORRECRSOTR_SOTRKEY);
					pTbl->SetValue(n,VOICERECR_SOTR_FIO,_GetFullFIO(sotrkey));
					if (sotrBossKey && pSotrList->IsKeyExist(sotrkey,&index))
					{
						sprintf(cfobuffer,"%09d",sotrBossKey[index]);
						pTbl->SetValue(n,VOICERECR_SOTR_BOSSCODE,cfobuffer);
						//pTbl->SetIntValue(n,VOICERECR_SOTR_BOSSCODE,sotrBossKey[index]);
					}
				}
			}
			if (gonorar!=0)
			{
				if(bIsFreeWorkGenDog)
				{
					FreeWorkAktAddInfo * addInfo = data->aktdata->GetFWAktAddData();
					if(addInfo)
					{ // #57287
						int numOtherCFO = 0;
						double otherCFOGonorar = 0;

						#if FW_GONORAR_NEW()
						sdword fullSumNet = 0;
						sdword fullSumNetWithNDS = 0;
						int mainCFOLine = -1;
						#endif //FW_GONORAR_NEW	

						const double realGonorar = DigitalInt2Double(addInfo->GetRealGonorar());
						needMask=ENUM2BIT(DOGOVORRECR_SORTFUNC_CFOBONUS); // сначала доп. цфо
						for (l=0;l<numline;l++)
						{
							if(!(pSotrQuary->GetIntField_MinNull(l,LIKEDOGOVORRECRSOTR_SOTRFUNC) & needMask))
								continue;

							percent = pSotrQuary->GetIntField(l,LIKEDOGOVORRECRSOTR_BONUSPERCENT);
							if (percent<=0) 
								continue;

							const int insLine=pTbl->GetNumLine();
							pTbl->InsertLine();
							pTbl->SetValue(insLine,VOICERECR_SOTR_FUNC, LS(IDS_05491));
							cfokey=pSotrQuary->GetIntField(l,LIKEDOGOVORRECRSOTR_BONUSCFO);
							pTbl->SetValue(insLine,VOICERECR_SOTR_FIO,_GetCFOName(cfokey));// спецом!
							sprintf(cfobuffer,"%09d",cfokey);
							pTbl->SetValue(insLine,VOICERECR_SOTR_CFOCODE,cfobuffer);
								
							gon=realGonorar*DigitalInt2Double(percent)/100;
							otherCFOGonorar += gon;
							gonwithnds=gon*(1+DigitalInt2Double(GetNDSPercentDigitalInt(data->aktdata->aktdate))/100.0);

							#if FW_GONORAR_NEW()
							fullSumNet += Double2DigitalInt(gon);
							fullSumNetWithNDS += Double2DigitalInt(gonwithnds);
							#endif //FW_GONORAR_NEW	

							pTbl->SetValue(insLine,VOICERECR_SOTR_SUM_NET,Double2Str(buf,gon,DOUBLETYPE_MONEY));
							sprintf(buffer,"%s%c",GetNDSPercentText(data->aktdata->aktdate),'%');
							pTbl->SetValue(insLine,VOICERECR_SOTR_SUM_NDSVAL,buffer);
							pTbl->SetValue(insLine,VOICERECR_SOTR_SUM_WITHNDS,Double2Str(buf,gonwithnds,DOUBLETYPE_MONEY));
							numOtherCFO++;
						}

						needMask=ENUM2BIT(DOGOVORRECR_SORTFUNC_MAIN_CFOBONUS); // теперь основное цфо - получает что осталось

						#if FW_GONORAR_NEW()
						sdword persentSum = 0;
						{
							KeyList objkeyList; objkeyList<<needMask;
							DPtr<UnsortedKeyList> allMainPercentsKl = MakeUnsortedKeyListFromSelectQuaryWithObjKeyList(pSotrQuary, LIKEDOGOVORRECRSOTR_BONUSPERCENT, &objkeyList, LIKEDOGOVORRECRSOTR_SOTRFUNC);
							for(keylistindex_t i=0; i<allMainPercentsKl->GetNumKey(); i++)
								persentSum += allMainPercentsKl->GetKey(i);
						}

						if(persentSum>0)
						#endif //FW_GONORAR_NEW		
						{ // if(persentSum>0)

							for (l=0;l<numline;l++)
							{
								if(!(pSotrQuary->GetIntField_MinNull(l,LIKEDOGOVORRECRSOTR_SOTRFUNC) & needMask))
									continue;

								percent = pSotrQuary->GetIntField(l,LIKEDOGOVORRECRSOTR_BONUSPERCENT);
								if (percent<=0) 
									continue;

								const int insLine = pTbl->GetNumLine() - numOtherCFO;
								pTbl->InsertLine(-1, insLine);
								pTbl->SetValue(insLine,VOICERECR_SOTR_FUNC, LS(IDS_05491));
								cfokey=pSotrQuary->GetIntField(l,LIKEDOGOVORRECRSOTR_BONUSCFO);
								pTbl->SetValue(insLine,VOICERECR_SOTR_FIO,_GetCFOName(cfokey));// спецом!
								sprintf(cfobuffer,"%09d",cfokey);
								pTbl->SetValue(insLine,VOICERECR_SOTR_CFOCODE,cfobuffer);

								gon = gonorar - otherCFOGonorar;

								#if FW_GONORAR_NEW()
								if(mainCFOLine == -1)
									mainCFOLine = insLine;

								gon = (gon/DigitalInt2Double(persentSum))*DigitalInt2Double(percent);
								#endif //FW_GONORAR_NEW		

								gonwithnds=gon*(1+DigitalInt2Double(GetNDSPercentDigitalInt(data->aktdata->aktdate))/100.0);

								#if FW_GONORAR_NEW()
								fullSumNet += Double2DigitalInt(gon);
								fullSumNetWithNDS += Double2DigitalInt(gonwithnds);
								#endif //FW_GONORAR_NEW	

								pTbl->SetValue(insLine,VOICERECR_SOTR_SUM_NET,Double2Str(buf,gon,DOUBLETYPE_MONEY));
								sprintf(buffer,"%s%c",GetNDSPercentText(data->aktdata->aktdate),'%');
								pTbl->SetValue(insLine,VOICERECR_SOTR_SUM_NDSVAL,buffer);
								pTbl->SetValue(insLine,VOICERECR_SOTR_SUM_WITHNDS,Double2Str(buf,gonwithnds,DOUBLETYPE_MONEY));
							}

							#if FW_GONORAR_NEW()
							// Если из-за округления осталась разница, то добавляем ее к основному ЦФО
							AdjustTblLineValueByDelta(pTbl, mainCFOLine, VOICERECR_SOTR_SUM_NET, data->gonorarsum - fullSumNet);
							AdjustTblLineValueByDelta(pTbl, mainCFOLine, VOICERECR_SOTR_SUM_WITHNDS, data->voicesum - fullSumNetWithNDS);
							#endif //FW_GONORAR_NEW	

						} // if(persentSum>0)
					}
				}
				else
				{
					needMask=ENUM2BIT(DOGOVORRECR_SORTFUNC_BONUS);
					for (l=0;l<numline;l++)
					{
						if (pSotrQuary->GetIntField_MinNull(l,LIKEDOGOVORRECRSOTR_SOTRFUNC) & needMask)
						{
							percent = pSotrQuary->GetIntField(l,LIKEDOGOVORRECRSOTR_BONUSPERCENT);
							if (percent<=0)
								continue;

							n=pTbl->GetNumLine();
							pTbl->InsertLine();
							pTbl->SetValue(n,VOICERECR_SOTR_FUNC,LS(IDS_05491));
							sotrkey=pSotrQuary->GetIntField(l,LIKEDOGOVORRECRSOTR_SOTRKEY);
							cfokey=pSotrQuary->GetIntField(l,LIKEDOGOVORRECRSOTR_BONUSCFO);
							pTbl->SetValue(n,VOICERECR_SOTR_FIO,_GetFullFIO(sotrkey));
							pTbl->SetValue(n,VOICERECR_SOTR_CFO,_GetCFOName(cfokey));
							sdword numpplpersotr = pSotrQuary->GetIntField(l,LIKEDOGOVORRECRSOTR_NUMPPLPERSOTR);
							pTbl->SetValue(n,VOICERECR_SOTR_NUMPPLPERSOTR,DigitalInt2Text(numpplpersotr));
							
							sprintf(cfobuffer,"%09d",cfokey);
							pTbl->SetValue(n,VOICERECR_SOTR_CFOCODE,cfobuffer);
							//pTbl->SetIntValue(n,VOICERECR_SOTR_CFOCODE,cfokey);
							if (sotrBossKey && pSotrList->IsKeyExist(sotrkey,&index))
							{
								sprintf(cfobuffer,"%09d",sotrBossKey[index]);
								pTbl->SetValue(n,VOICERECR_SOTR_BOSSCODE,cfobuffer);
								//pTbl->SetIntValue(n,VOICERECR_SOTR_BOSSCODE,sotrBossKey[index]);
							}
							gon=gonorar*DigitalInt2Double(percent)/100;
							gonwithnds=gon*(1+DigitalInt2Double(GetNDSPercentDigitalInt(data->aktdata->aktdate, data->aktdata->dogdata->residentType))/100.0);
							sprintf(buffer,LS(IDS_01295),Double2Str(buf,gon,DOUBLETYPE_MONEY),Double2Str(buf1,gonwithnds,DOUBLETYPE_MONEY));
							pTbl->SetValue(n,VOICERECR_SOTR_SUM,buffer);
							pTbl->SetValue(n,VOICERECR_SOTR_SUM_NET,Double2Str(buf,gon,DOUBLETYPE_MONEY));
							sprintf(buffer,"%s%c",GetNDSPercentText(data->aktdata->aktdate, data->aktdata->dogdata->residentType),'%');
							pTbl->SetValue(n,VOICERECR_SOTR_SUM_NDSVAL,buffer);
							pTbl->SetValue(n,VOICERECR_SOTR_SUM_WITHNDS,Double2Str(buf,gonwithnds,DOUBLETYPE_MONEY));
						}
					}
				}
			}
		}
	}
	DELETE(pCFOList);
	DELETE(pOtvCFOList);
	DELETE(pSotrList);
	FREE_(sotrBossKey);
	SetDateDelimitor(0);
	return pData;
}
/*
void MakeVoiceRecrZayavka(HWND hWnd,sdword voicekey,int printOnly)
{
	char* errStr=0;
	int wordTmplID=WORDTMPL_VOICEZAYAVKA_REKR;
	VoiceRecrData* data=LoadVoiceRecrCardData(hWnd,voicekey);
	SelectQuary* pSotrQuary=DUPObjSelectQuary(hWnd,OBJTYPE_VOICERECR,voicekey,VOICERECROBJ_SOTR);
	if (data && pSotrQuary)
	{
		DataBaseData* pData=MakeVoiceRecrZayavkaWinWordData(hWnd,data,pSotrQuary);
		errStr=MakeDocFromTmpl(hWnd,GetWordDocTmpl(wordTmplID,0),
		//errStr=MakeDocFromTmpl(hWnd,GetWordDocTmpl(hWnd,data->aktdata->dogdata->ancorkey,data->aktdata->dogdata->ancorfilial,wordTmplID,0),
							pData,printOnly?1:0);
		DELETE(pData);
	}
	DELETE(data);
	DELETE(pSotrQuary);
	StdProceedError(hWnd,errStr);
}
*/
DataBaseData* MakeDataForVoiceRecrZayavka(HWND hWnd,sdword voicekey,sdword* ancorkadryorgkey,sdword* ancorfilial)
{
	DataBaseData* pData=0;
	VoiceRecrData* data=LoadVoiceRecrCardData(hWnd,voicekey);
	SelectQuary* pSotrQuary=DUPObjSelectQuary(hWnd,OBJTYPE_VOICERECR,voicekey,VOICERECROBJ_SOTR);
	if (data && pSotrQuary)
	{
		pData=MakeVoiceRecrZayavkaWinWordData(hWnd,data,pSotrQuary);
		if (ancorkadryorgkey) *ancorkadryorgkey=data->aktdata->dogdata->ancorkey;
		if (ancorfilial) *ancorfilial=data->aktdata->dogdata->ancorfilial;
	}
	DELETE(data);
	DELETE(pSotrQuary);
	return pData;
}
void MakeVoiceRecrZayavka(HWND hWnd,sdword voicekey,int printOnly)
{
	sdword ancorkadryorgkey=0,ancorfilial=0;
	DataBaseData* pData=MakeDataForVoiceRecrZayavka(hWnd,voicekey,&ancorkadryorgkey,&ancorfilial);
	if (pData)
	{
		int wordTmplID=WORDTMPL_VOICEZAYAVKA_REKR;
		char* errStr=MakeDocFromTmpl(hWnd,GetWordDocTmpl(wordTmplID,0),
		//char* errStr=MakeDocFromTmpl(hWnd,GetWordDocTmpl(hWnd,ancorkadryorgkey,ancorfilial,wordTmplID,0),
							pData,printOnly?1:0);
		DELETE(pData);
		StdProceedError(hWnd,errStr);
	}
}

DogovorRecrData* LoadDogovorRecrCardData(HWND hWnd,sdword dogkey);
const char* EnsiKey2EnsiKeyStr(sdword ensikey);

bool MakeVoiceRecrZayavkaInExcel(HWND hWnd, sdword voicekey, bool bPrintOnly, bool bShowSavedTmpFile)
{
	bool retVal = false;
	if(!IsMSOAvailable(MSO_EXCEL))
	{
		ErrorBox(hWnd, "Не найден MS Excel!");
		return retVal;
	}

	DataBaseData* pData=MakeDataForVoiceRecrZayavka(hWnd,voicekey,0,0);
	VoiceRecrAddData* addData=GetVoiceRecrAddData(voicekey);
	const VoiceRecrData* data=addData?addData->GetVoiceRecrData():0;

	int numLines = 0, numRows = 9;
	if (pData && data)
	{
		// SKOREX const bool isSkorexDog = ISMAINDOGOVORRECRTYPE_SKOREX_KONTRAKT(data->GetDogovorType());
		const bool isFreeWorkGenDog = (data->GetDogovorType() == DOGOVORRECRTYPE_LEASING_GEN);
		FreeWorkAktAddInfo * addInfo = data->aktdata ? data->aktdata->GetFWAktAddData() : 0;

		TblClipBuffer clip;
		clip.PassCell("Заявка на выставление счета",0,1); 
		clip.PassCell("",0,1);clip.PassCell("",0,1);clip.PassCell("",0,1);
		clip.PassCell("ИНН",0,1); clip.PassCell("КПП",0,1);
		clip.PassCell("Код МЕТА",0,1);clip.PassCell("Код ЕНСИ ЦФО",0,1);clip.PassCell("Код ЕНСИ сотрудника",0,1); 
		clip.PassCell("Код ЕНСИ договора",0,1);clip.PassCell("Код ЕНСИ контрагента",1,1); 

		numLines++;
		//clip.PassLine("",0);
//#define PASSLINEWITHPROMPT(pr,fldid){clip.PassCell(pr,0,1);clip.PassCell(pData->GetValue(fldid),1,1);}
		clip.PassCell("Дата акта",0,1);clip.PassCell(_db2viewdate(data->aktdata->aktdate),1,1); numLines++;
		clip.PassCell("Основание",0,1);clip.PassCell(pData->GetValue(VOICERECR_MAIN_VOICE_TYPE),1,1); numLines++;
		clip.PassCell("Гонорар за услугу (по акту)",0,1);clip.PassCell(pData->GetValue(VOICERECR_MAIN_AKT_SERVICEFEE),1,1); numLines++;
		clip.PassCell("Общая сумма акта (с НДС)",0,1);clip.PassCell(pData->GetValue(VOICERECR_MAIN_AKT_SUM),1,1); numLines++;

		{
			clip.PassCell("Сумма счета (с НДС)",0,1);
			if(isFreeWorkGenDog) 
			{
				clip.PassCell(pData->GetValue(VOICERECR_MAIN_VOICE_SUM),0,1); 
				char nds[10];
				_snprintf_s(nds, SIZEOFARRAY(nds), _TRUNCATE, "%s%%", GetNDSPercentText(data->aktdata->aktdate));
				clip.PassCell(nds,1,1);
			}
			else
			{
				clip.PassCell(pData->GetValue(VOICERECR_MAIN_VOICE_SUM),1,1); 
			}
			numLines++;
		}

		if(isFreeWorkGenDog)
		{
			clip.PassCell("Сумма счета (без НДС)",0,1);clip.PassCell(pData->GetValue(VOICERECR_MAIN_VOICE_GONORAR_SUM),1,1); numLines++;
			if(addInfo)
			{
				clip.PassCell(SetMyCompanyNameStrInStr(LS(IDS_10991)),0,1); 
				clip.PassCell(DigitalInt2Text(addInfo->GetRealGonorar()),1,1); numLines++;
			}
		}
		
		if(!isFreeWorkGenDog)
		{
			clip.PassCell("Тип услуги",0,1); clip.PassCell(pData->GetValue(VOICERECR_MAIN_USLUGA_TYPE),1,1); numLines++;
		}

		if (data->classifCodeList.GetNumKey())
		{
			const sdword code = data->classifCodeList.GetKey(0);
			char code1C[255]; *code1C = 0;
			const char* vocabName = ClassVocabCode2ClassVocabName(CLASSTYPE_VOICEUSLUGA, code);
			const char* name1C = BreakClassifFromFullMLStr(code1C, vocabName, SIZEOFARRAY(code1C));
			clip.PassCell("Услуга",0,1);
			clip.PassCell(name1C,0,1);
			clip.PassCell(code1C,1,1); numLines++;
		}

		if (isFreeWorkGenDog)
		{
			clip.PassCell("Руководитель ЦФО",0,1); clip.PassCell(pData->GetValue(VOICERECR_MAIN_CFO_CHIEF),1,1); numLines++;
			clip.PassCell("Подписант",0,1); clip.PassCell(_GetFullFIO(data->aktdata->whosign_ancor),1,1); numLines++;
		}
		else
		{
			clip.PassCell("Ниша (отрасль)",0,1);clip.PassCell(pData->GetValue(VOICERECR_MAIN_CFO_WHATDOCOMPANY),1,1); numLines++;

			clip.PassCell("Руководитель ЦФО",0,1);clip.PassCell(pData->GetValue(VOICERECR_MAIN_CFO_CHIEF),0,1);
			clip.PassCell("",0,1);clip.PassCell("",0,1);
			clip.PassCell("",0,1);clip.PassCell("",0,1);
			clip.PassCell("",0,1);clip.PassCell("",0,1);
			clip.PassCell(pData->GetValue(VOICERECR_MAIN_CFO_CHIEFCODE),1,1); numLines++;
		}

		clip.PassCell((isFreeWorkGenDog /* SKOREX || isSkorexDog*/)?"Юр.лицо исполнителя":"Юр.лицо Анкора",0,1);clip.PassCell(pData->GetValue(VOICERECR_MAIN_ANCORNAME),0,1);
			clip.PassCell("",0,1);clip.PassCell("",0,1);
			clip.PassCell(pData->GetValue(VOICERECR_MAIN_ANCORKADRY_INN),0,1);
			clip.PassCell(pData->GetValue(VOICERECR_MAIN_ANCORKADRY_KPP),0,1);
			clip.PassCell(pData->GetValue(VOICERECR_MAIN_ANCORKADRYCODE),1,1); numLines++;
		clip.PassCell("Юр.лицо заказчика",0,1);clip.PassCell(pData->GetValue(VOICERECR_MAIN_ZAKAZNAME),0,1);
			clip.PassCell("",0,1);clip.PassCell("",0,1);
			clip.PassCell(pData->GetValue(VOICERECR_MAIN_ZAKAZKADRY_INN),0,1);
			clip.PassCell(pData->GetValue(VOICERECR_MAIN_ZAKAZKADRY_KPP),0,1);
			clip.PassCell(pData->GetValue(VOICERECR_MAIN_ZAKAZKADRYCODE),0,1); 
			clip.PassCell("",0,1);clip.PassCell("",0,1);clip.PassCell("",0,1);
			clip.PassCell(data->aktdata->dogdata->zakazcontragentensikey,1,1);
			numLines++;

		if (isFreeWorkGenDog)
		{
			if(addInfo)
			{
				clip.PassCell("Услуга",0,1);
					clip.PassCell(addInfo->GetStrForVoice(),1,1);
					numLines++;
			}
		}
		else
		{
			char usluga[128] = {0}, buf[256] = {0};
			MakeUslugaStrForGenDogovor(usluga,DOGOVORLANG_1(data->aktdata->dwlang),data->aktdata->dogdata->poslevel);
			_snprintf_s(buf, SIZEOFARRAY(buf), _TRUNCATE, "Услуги %s на позицию %s", usluga, pData->GetValue(VOICERECR_MAIN_POSNAME));
			clip.PassCell("Позиция",0,1);clip.PassCell(buf,1,1); numLines++;
		}
		
		clip.PassLine("",0); numLines++;
		if (!isFreeWorkGenDog)
		{
			ListContainer* resultPpl = addData->GetResultPplData();
			const int numline = resultPpl->GetNumLine();
			for (int l=0; l < numline; l++)
			{
				if(l==0)
					clip.PassCell("Результат",0,1);

				clip.PassCell(_GetFullFIO(resultPpl->GetKey(l)),1,1); numLines++;
			}
		}

		char key[10]; 
		const int type = data->GetDogovorType();
		if(type == DOGOVORRECRTYPE_PROTOCOL || type == DOGOVORRECRTYPE_ADDKONTRAKT_EX)
		{
			DPtr<DogovorRecrData> upDogData = LoadDogovorRecrCardData(hWnd,data->aktdata->dogdata->updogovorkey);
			clip.PassCell("Ген. договор/Разовый договор",0,1);clip.PassCell(GetDogovorRecrDASName(upDogData,2),0,1);
				clip.PassCell("",0,1);clip.PassCell("",0,1);clip.PassCell("",0,1);clip.PassCell("",0,1);
				_itoa_s(data->aktdata->dogdata->updogovorkey,key,SIZEOFARRAY(key),10);
			clip.PassCell(key,0,1); 
			clip.PassCell("",0,1);clip.PassCell("",0,1);
			clip.PassCell(EnsiKey2EnsiKeyStr(data->aktdata->dogdata->updogovorensikey),1,1);
			numLines++;

			clip.PassCell("Протокол/Дополнительное соглашение",0,1);clip.PassCell(GetDogovorRecrDASName(data->aktdata->dogdata,2),0,1);
				clip.PassCell("",0,1);clip.PassCell("",0,1);clip.PassCell("",0,1);clip.PassCell("",0,1);
				_itoa_s(data->aktdata->dogkey,key,SIZEOFARRAY(key),10);
			clip.PassCell(key,0,1);
			clip.PassCell("",0,1);clip.PassCell("",0,1);
			clip.PassCell(EnsiKey2EnsiKeyStr(data->aktdata->dogdata->dogovorensikey),1,1);
			numLines++;
		}
		else
		{
			clip.PassCell("Ген. договор/Разовый договор",0,1);clip.PassCell(GetDogovorRecrDASName(data->aktdata->dogdata,2),0,1);
			clip.PassCell("",0,1);clip.PassCell("",0,1);clip.PassCell("",0,1);clip.PassCell("",0,1);
			_itoa_s(data->aktdata->dogkey,key,SIZEOFARRAY(key),10);
			clip.PassCell(key,0,1);
			clip.PassCell("",0,1);clip.PassCell("",0,1);
			clip.PassCell(EnsiKey2EnsiKeyStr(data->aktdata->dogdata->dogovorensikey),1,1);
			numLines++;

			if (!isFreeWorkGenDog)
			{
				clip.PassCell("Протокол",0,1);clip.PassCell("",1,1); numLines++;
			}
		}
		
		clip.PassCell("Комментарий",0,1);clip.PassCell(pData->GetValue(VOICERECR_MAIN_COMMENT),1,1); numLines++;

		//PASSLINEWITHPROMPT("Основание",VOICERECR_MAIN_VOICE_TYPE);
		//PASSLINEWITHPROMPT("Общая сумма акта (с НДС)",VOICERECR_MAIN_AKT_SUM);
		//PASSLINEWITHPROMPT("Сумма счета (с НДС)",VOICERECR_MAIN_VOICE_SUM);
		//PASSLINEWITHPROMPT("Тип услуги",VOICERECR_MAIN_USLUGA_TYPE);
		//PASSLINEWITHPROMPT("Ниша (отрасль)",VOICERECR_MAIN_CFO_WHATDOCOMPANY);	
		//PASSLINEWITHPROMPT("Руководитель ЦФО",VOICERECR_MAIN_CFO_CHIEF);	
		//PASSLINEWITHPROMPT("Юр.лицо Анкора",VOICERECR_MAIN_ANCORNAME);
		//PASSLINEWITHPROMPT("Юр.лицо заказчика",VOICERECR_MAIN_ZAKAZNAME);
		//PASSLINEWITHPROMPT("Позиция",VOICERECR_MAIN_POSNAME);
		//PASSLINEWITHPROMPT("Комментарий",VOICERECR_MAIN_COMMENT);
//#undef PASSLINEWITHPROMPT
		clip.PassLine("",0); numLines++;
		clip.PassLine("Сотрудники",0); numLines++;

		TblData* pTbl=pData->GetTblData(TBL_VOICERECR_SOTR);
		const int numline=pTbl->GetNumLine();
		for (int l=0;l<numline;l++)
		{
			clip.PassCell(pTbl->GetViewValue(l,VOICERECR_SOTR_FUNC),0,1);
			clip.PassCell(pTbl->GetViewValue(l,VOICERECR_SOTR_FIO),0,1);
			clip.PassCell(pTbl->GetViewValue(l,VOICERECR_SOTR_CFO),0,1);
			//clip.PassCell(pTbl->GetViewValue(l,VOICERECR_SOTR_SUM),1,1);
			clip.PassCell(pTbl->GetViewValue(l,VOICERECR_SOTR_SUM_NET),0,1);
			clip.PassCell(pTbl->GetViewValue(l,VOICERECR_SOTR_SUM_NDSVAL),0,1);
			clip.PassCell(pTbl->GetViewValue(l,VOICERECR_SOTR_SUM_WITHNDS),0,1);
			clip.PassCell(pTbl->GetViewValue(l,VOICERECR_SOTR_NUMPPLPERSOTR),0,1);
			clip.PassCell(pTbl->GetViewValue(l,VOICERECR_SOTR_CFOCODE),0,1);
			clip.PassCell(pTbl->GetViewValue(l,VOICERECR_SOTR_BOSSCODE),1,1); numLines++;
		}
		clip.PassLine("",1);

		DELETE(pData);

		HGLOBAL hClip=clip.EndTask();
		if (hClip)
		{
			if (!OpenClipboard(hMainWnd))
			{
				GlobalFree(hClip);
			}
			else
			{
				EmptyClipboard();
				SetClipboardLocale();
				if( !SetClipboardData(CF_TEXT,hClip) )
					GlobalFree(hClip);
				CloseClipboard();

				//MakeExcelDoc(hWnd,0,1);

				ExcelExportCellsStyle cellsStyle;
				cellsStyle.style.Init();
				cellsStyle.style.fontName = "Arial Cyr";
				cellsStyle.style.fontSize = 10;
				cellsStyle.style.textWrap = TRUE;
				cellsStyle.style.verAlign = Excel::xlTop;
				cellsStyle.style.horAlign = Excel::xlLeft;
				cellsStyle.style.cellFormat = "@";

				// ширина колонок указывается эмпирическим путем.
				// Для Excel пересчитывается обратно в char ( ExcelAutomation::FromPx2Chars )
				const int basicWidth = 180; //  для шрифта Arial Cyr 10
				cellsStyle.colWidth.AddKey( basicWidth );
				cellsStyle.colWidth.AddKey( basicWidth );
				cellsStyle.colWidth.AddKey( basicWidth * 8 / 10 );
				cellsStyle.colWidth.AddKey( basicWidth / 2 );
				cellsStyle.colWidth.AddKey( basicWidth / 2);
				cellsStyle.colWidth.AddKey( basicWidth / 2 );
				cellsStyle.colWidth.AddKey( -1 ); // default
				cellsStyle.colWidth.AddKey( basicWidth / 3 );
				cellsStyle.colWidth.AddKey( basicWidth / 3 );


				ExcelExportParams exportParams;
				exportParams.hParentWnd = hWnd;
				exportParams.bBoldCellBordes = true;
				exportParams.bFitOnePage = true;
				exportParams.bShowSavedFile = bShowSavedTmpFile;
				exportParams.sheetName = LS(IDS_03345);
				exportParams.cs = &cellsStyle;
				exportParams.szTable.cx = numRows;
				exportParams.szTable.cy = numLines;

				char tempFolderPath[_MAX_PATH];
				GetTempPath(_MAX_PATH, tempFolderPath);
				GetTempFileName(tempFolderPath, "Exp", 0, exportParams.xlsFileName);

				EnumExcelExportVerb xev = bPrintOnly ? xevPrintLandscape : xevExport;
				
				HRESULT hr = ClipBoard2Excel(xev, exportParams);
				if( FAILED(hr) )
					return retVal;

				if( !bPrintOnly )
				{
					char * errStr = AddBinaryDocToObjLikeDogovor(hWnd,
																OBJTYPE_VOICERECR,voicekey,exportParams.xlsFileName,
																VOICERECR_DOCTYPE_VOICEREQUEST_EXCEL,
																GetFileExt(exportParams.xlsFileName),
																LS(IDS_06800));
					if(!StdProceedError(hWnd, errStr))
						return retVal;
				}

				return true;
			}
		}
	}

	return retVal;
}
