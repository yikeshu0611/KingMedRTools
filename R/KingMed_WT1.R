#' To clear WT1 data from King Med test to a dataframe
#'
#' @param filename the name of fusion gene data text
#' @param WriteExecl A logical object. TRUE as default to write excel
#'
#' @return a dataframe
#' @export
#'
#' @examples KingMed_WT1("data.txt")
KingMed_WT1<-function(filename,WriteExecl=TRUE){
  {
  library(textreadr)
  library(stringr)
  library(stringi)
  library(magrittr)
  library(plyr)
  library(tmcn)
  }
  #read the long string
  if (grepl(".txt",filename)) { TRdata=suppressWarnings(toString(readLines(filename)))
  }else if(grepl(".doc",filename) | grepl(".docx",filename)) {TRdata=toString(read_docx(filename))}
  #standardize characters
  TRdata=stri_trans_nfkd(TRdata)
  TRdata=TRdata %>%
    gsub(pattern = toUTF8("\u533B \u9662"),replacement = toUTF8("\u533B\u9662"))%>%
    gsub(pattern = toUTF8("\u79D1 \u5BA4"),replacement = toUTF8("\u79D1\u5BA4"))%>%
    gsub(pattern = toUTF8("\u6027 \u522B"),replacement = toUTF8("\u6027\u522B"))%>%
    gsub(pattern = toUTF8("\u5E74 \u9F84"),replacement = toUTF8("\u5E74\u9F84"))%>%
    gsub(pattern = toUTF8("CD45\u5F31, \u8868\u8FBE\u7EC6\u80DE"),replacement = toUTF8("CD45\u5F31\u8868\u8FBE\u7EC6\u80DE"))%>%
    gsub(pattern = toUTF8("CD45\u9634\u6027, \u8868\u8FBE\u7EC6\u80DE"),replacement = toUTF8("CD45\u9634\u6027\u8868\u8FBE\u7EC6\u80DE"))%>%
    gsub(pattern = toUTF8("\u5355\u6838\u7EC6, \u80DE"),replacement = toUTF8("\u5355\u6838\u7EC6\u80DE"))%>%
    gsub(pattern = toUTF8("\u7EFC, \u5408\u8003\u8651"),replacement = toUTF8("\u7EFC\u5408\u8003\u8651"))%>%
    gsub(pattern = toUTF8("\u5F62, \u6001\u5B66"),replacement = toUTF8("\u5F62\u6001\u5B66"))%>%
    gsub(pattern = toUTF8("\u5176, \u514D\u75AB\u8868\u578B"),replacement = toUTF8("\u5176\u514D\u75AB\u8868\u578B"))%>%
    gsub(pattern = toUTF8("\u8F85, \u52A9\u8BCA\u65AD\u65B9\u6CD5"),replacement = toUTF8("\u8F85\u52A9\u8BCA\u65AD\u65B9\u6CD5"))%>%
    gsub(pattern = toUTF8("\u5176\u5B83, \u4E34\u5E8A\u4FE1\u606F"),replacement = toUTF8("\u5176\u5B83\u4E34\u5E8A\u4FE1\u606F"))%>%
    gsub(pattern = toUTF8("\u59D3 \u540D"),replacement = toUTF8("\u59D3\u540D"))%>%
    gsub(pattern = toUTF8(" \u8F7B\u94FE"),replacement = toUTF8("\u8F7B\u94FE"))%>%
    gsub(pattern = toUTF8(" \u6E38\u79BB\u8F7B\u94FE"),replacement = toUTF8("\u6E38\u79BB\u8F7B\u94FE"))%>%
    gsub(pattern = toUTF8("ABL1 \u57FA\u56E0\u62F7\u8D1D\u6570"),replacement = toUTF8("ABL1\u57FA\u56E0\u62F7\u8D1D\u6570"))

  #split
  tongren=toUTF8('\u4E0A\u6D77\u4EA4\u901A\u5927\u5B66\u533B\u5B66\u9662\u9644\u5C5E\u540C\u4EC1\u533B\u9662')
  TRdata1=strsplit(x = TRdata,split = tongren)[[1]]
  #get purpose data
  #abnormal immueglobal emia
  WT1=toUTF8("WT1\u57FA\u56E0\u5B9A\u91CF\u68C0\u6D4B")
  dataget=c()
  for (data.i in 1:length(TRdata1)) {
    if (grepl(WT1,TRdata1[data.i])){
      dataget=c(dataget,data.i)
    }
  }
  #extract abnormal ig data
  dataget
  length(dataget)
  TRdata2=abnormalig=TRdata1[dataget]
  #
  TR.datafrme=data.frame()
  prgbar<- txtProgressBar(min = 0, max = length(dataget),
                          style = 3,
                          initial = 0,width = 35)
  for (flui.i in 1:length(dataget)) {
    #page i
    TRdata2.i=TRdata2[flui.i]
    if (grepl(toUTF8("\u8FDF\u53D1\u9879\u76EE"),TRdata2.i)){next(flui.i)}
    #Title
    {
    #sample code
    SampleCode=toUTF8("\u6807\u672C\u6761\u7801")
    SampleCode.Num=gsub(pattern = " ",replacement = "",
                        x=gsub(pattern = SampleCode,replacement = "",
                               x =str_extract(string = TRdata2.i,
                                              pattern = paste0(SampleCode,"[0-9a-zA-Z ]{1,}"))))
    #PatientName
    PatientName=toUTF8("\u75C5\u4EBA\u59D3\u540D")
    PatientName.Num=gsub(pattern = " ",replacement = "",
                         x=sub(pattern = toUTF8("\u79D1\u5BA4"),replacement = "",
                               sub(pattern = toUTF8("\u59D3\u540D"),replacement = "",
                                   x =str_extract(string = TRdata2.i,
                                                  pattern = paste0(toUTF8("\u59D3\u540D"),".*",toUTF8("\u79D1\u5BA4"))))))
    #Department
    Department=toUTF8("\u79D1\u5BA4")
    Department.Num=gsub(pattern = " ",replacement = "",
                        x=sub(pattern = toUTF8("\u5B9E\u9A8C\u53F7"),replacement = "",
                              sub(pattern = Department,replacement = "",
                                  x =str_extract(string = TRdata2.i,
                                                 pattern = paste0(Department,".*",toUTF8("\u5B9E\u9A8C\u53F7"))))))
    #sex
    sex=toUTF8("\u6027\u522B")
    sex.Num=gsub(pattern = " ",replacement = "",
                 x=sub(pattern = toUTF8("\u623F/\u5E8A\u53F7"),replacement = "",
                       sub(pattern = sex,replacement = "",
                           x =str_extract(string = TRdata2.i,
                                          pattern = paste0(sex,".*",toUTF8("\u623F/\u5E8A\u53F7"))))))
    #bedid
    bedid=toUTF8("\u623F/\u5E8A\u53F7")
    bedid.Num=gsub(pattern = " ",replacement = "",
                   x=sub(pattern = toUTF8("\u95E8\u8BCA/\u4F4F\u9662\u53F7"),replacement = "",
                         sub(pattern = bedid,replacement = "",
                             x =str_extract(string = TRdata2.i,
                                            pattern = paste0(bedid,".*",toUTF8("\u95E8\u8BCA/\u4F4F\u9662\u53F7"))))))
    if (substring(bedid.Num,1,1)=="+"){bedid.Num=sub(pattern = "\\+",replacement = "",
                                              x = paste0(toUTF8("\u4E34"),bedid.Num))}
    #Hospitalid
    Hospitalid=toUTF8("\u95E8\u8BCA/\u4F4F\u9662\u53F7")
    Hospitalid.Num=gsub(pattern = " ",replacement = "",
                        x=sub(pattern =  paste0('[ ,]*',toUTF8("\u5E74\u9F84")),replacement = "",
                              sub(pattern = toUTF8("\u95E8\u8BCA/\u4F4F\u9662\u53F7"),replacement = "",
                                  x =str_extract(string = TRdata2.i,
                                                 pattern = paste0(toUTF8("\u95E8\u8BCA/\u4F4F\u9662\u53F7"),".*", toUTF8("\u5E74\u9F84"))))))
    #age
    age=toUTF8("\u5E74\u9F84")
    age.Num=gsub(pattern = " ",replacement = "",
                 x=sub(pattern =  paste0('[ ,]*',toUTF8("\u7533\u8BF7\u533B\u751F")),replacement = "",
                       sub(pattern = age,replacement = "",
                           x =str_extract(string = TRdata2.i,
                                          pattern = paste0(age,".*", toUTF8("\u7533\u8BF7\u533B\u751F"))))))
    #ApplyDoctor
    ApplyDoctor=toUTF8("\u7533\u8BF7\u533B\u751F")
    ApplyDoctor.Num=sub(pattern = " ",replacement = "",
                        x=sub(pattern =  paste0('[ ,]*',toUTF8("\u63A5\u6536\u65E5\u671F")),replacement = "",
                              sub(pattern = ApplyDoctor,replacement = "",
                                  x =str_extract(string = TRdata2.i,
                                                 pattern = paste0(ApplyDoctor,".*", toUTF8("\u63A5\u6536\u65E5\u671F"))))))
    #bacterial
    bacterial=toUTF8("\u9001\u68C0\u6750\u6599")
    bacterial.Num=sub(pattern = " ",replacement = "",
                      x=sub(pattern =  paste0('[ ,]*',toUTF8("\u533B\u9662\u6807\u8BC6")),replacement = "",
                            sub(pattern = toUTF8("\u9001\u68C0\u6807\u672C"),replacement = "",
                                x =str_extract(string = TRdata2.i,
                                               pattern = paste0(toUTF8("\u9001\u68C0\u6807\u672C"),".*", toUTF8("\u533B\u9662\u6807\u8BC6"))))))
    #SamplingTime
    SamplingTime=toUTF8("\u91C7\u6837\u65F6\u95F4")
    SamplingTime.Num=gsub(pattern =  '[a-zA-Z ,]*',replacement = "",
                          sub(pattern = toUTF8("\u91C7\u6837\u65E5\u671F"),replacement = "",
                              x =str_extract(string = TRdata2.i,
                                             pattern = paste0(toUTF8("\u91C7\u6837\u65E5\u671F"),"[a-zA-Z0-9- ]*"))))
    #SamplingCondition
    SamplingCondition=toUTF8("\u6807\u672C\u60C5\u51B5")
    SamplingCondition.Num=gsub(pattern = " ",replacement = "",
                  x=sub(pattern =  toUTF8("\u8054\u7CFB\u7535\u8BDD"),replacement = "",
                              sub(pattern = SamplingCondition,replacement = "",
                                  x =str_extract(string = TRdata2.i,
                                                 pattern = paste0(SamplingCondition,".*",toUTF8("\u8054\u7CFB\u7535\u8BDD"))))))

    }
    {
    #get data frame
    first_title_location=c()
    last_title_location=c()
    first_title_location=max(gregexpr(WT1,TRdata2.i)[[1]][1])
    last_title_location=max(gregexpr(toUTF8("\u5EFA\u8BAE\u4E0E\u89E3\u91CA:"),TRdata2.i)[[1]][1])
    DataFrame=substring(TRdata2.i,first_title_location,last_title_location-1)%>%
      gsub(pattern = paste0(" ","\\S*",'[toUTF8("\u6cd5")(PCR)]'," "),replacement = ";")%>%
      gsub(pattern = " ",replacement = ";")%>%
      sub(pattern = WT1,replacement = "")%>%
      sub(pattern = toUTF8("\u8840\u6E05\u86CB\u767D\u5B9A\u91CF\u7EC4\u5408"),replacement = "")%>%
      gsub(pattern = ",;",replacement = ",")%>%
      gsub(pattern = ";%",replacement = "%")%>%
      gsub(pattern = ",{0,1}fr_flag,{0,1}",replacement = "")%>%
      sub(pattern = paste0("SP;G;A;M;",".*"),replacement = "")%>%
      sub(pattern = "\\(RQ-PCR\\)",replacement = "")
    DataFramelist=strsplit(x = DataFrame,split = ",")
    #in order to delet space list unit
    deletn=c()
    for (j in 1:length(DataFramelist[[1]])) {
        if (nchar(DataFramelist[[1]][j])==0){
          deletn=c(deletn,j)
        }
    }
    if (length(deletn)>=1){
      DataFramelist=DataFramelist[[1]][-deletn]
    }
    DataFramelist2=strsplit(x = DataFramelist,split = ";")
    Results=c()
    for (fgunlist.i in DataFramelist2) {
      fgunlist.clear= fgunlist.i[2] %>%
        sub(pattern = ".*\\(",replacement = "")%>%
        sub(pattern = "\\)",replacement = "")
      Results=c(Results,fgunlist.clear)
    }
    Results=ifelse(Results=="+",toUTF8("\u9633\u6027"),Results)
    Iterm=c()
    for (fgunlist.i in DataFramelist2) {
      Iterm=c(Iterm,fgunlist.i[1])
    }

    fgfinaledata=data.frame(matrix(data = Results,nrow = 1))
    colnames(fgfinaledata)=Iterm

  }
    ##########
    #dataframe
    TR.datafrme.i=c()
    TR.datafrme.i=cbind(SampleCode.Num,
                        PatientName.Num,
                        Department.Num,
                        sex.Num,
                        bedid.Num,
                        Hospitalid.Num,
                        age.Num,
                        ApplyDoctor.Num,
                        SamplingTime.Num,
                        SamplingCondition.Num)
    colnames(TR.datafrme.i)= c(SampleCode,
                               PatientName,
                               Department,
                               sex,
                               bedid,
                               Hospitalid,
                               age,
                               ApplyDoctor,
                               SamplingTime,
                               SamplingCondition)
    TR.datafrme.bind=cbind(TR.datafrme.i,fgfinaledata)
    TR.datafrme=rbind.fill(TR.datafrme,TR.datafrme.bind)
    setTxtProgressBar(pb = prgbar, value = flui.i)
  }
  close(prgbar)
  TR.datafrme1=TR.datafrme
  TR.datafrme=data.frame(ifelse(is.na(as.matrix(TR.datafrme1)),"",as.matrix(TR.datafrme1)))
  colnames(TR.datafrme)=colnames(TR.datafrme1)
  cat( toUTF8("\u68C0\u6D4B\u9879\u76EE:"),
       "\n",
    toUTF8("WT1\u57FA\u56E0\u5B9A\u91CF\u68C0\u6D4B"))
  if (WriteExecl==TRUE){
    filenamewritten=Sys.time() %>%
      sub(pattern = ":",replacement = toUTF8("\u02B1"))%>%
      sub(pattern = ":",replacement = toUTF8("\u5206"))%>%
      paste0(toUTF8("\u79D2 WT1\u57FA\u56E0\u5B9A\u91CF\u68C0\u6D4B.csv"))
    write.csv(TR.datafrme,filenamewritten)
    cat("\n","\n",
      toUTF8("\u6570\u636E\u88AB\u5199\u5165Excel\u4E2D,\u6587\u4EF6\u540D\u4E3A:"),
        "\n",
        filenamewritten)
  }
  return(TR.datafrme)
}
