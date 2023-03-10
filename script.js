
export function numToString(number){
    if (isNaN(number)){return undefined};

    const quydoi = {"0": "không", "1": "một","2":"hai", "3":"ba", "4":"bốn", "5":"năm", "6":"sáu", "7":"bảy", "8":"tám", "9":"chín"}
  
    let splitNumber = [...number.toString()];
    for (let i =0;splitNumber.length-1;i++){
      if(splitNumber[0]!=="0"){break;}
      splitNumber.shift();
    }
    if(splitNumber.length % 3 ===1){
        splitNumber.unshift("0","0")
    }else if(splitNumber.length % 3 ===2){
        splitNumber.unshift("0")
    };
    const groupNumber = [];
    for (let i=splitNumber.length/3;i!==0;i--){
        let tempArray = splitNumber.slice(0,3);
        splitNumber.splice(0,3);
        groupNumber.push(tempArray);
    }
    
    const groupChu = groupNumber.map((miniGroup, index)=>{
        let iLetter = "";
        switch(miniGroup[2]){
            case "0":break;
            case "5": 
                switch(miniGroup[1]){
                    case "0": iLetter="năm";break;
                    default: iLetter="lăm";break;
                }; break;
            case "1":
                switch(miniGroup[1]){
                    case "0": iLetter="một";break;
                    case "1": iLetter="một";break;
                    default: iLetter="mốt";break;
                }; break;
            default: iLetter=quydoi[miniGroup[2]];break;
        }
        switch(miniGroup[1]){
            case "0":
                if(index===0 && miniGroup[0]==="0"){break;}
                if(miniGroup[2]!=="0"){
                    iLetter= `lẻ ${iLetter}`;
                };
            break;
            case "1": iLetter = `mười ${iLetter}`; break;
            default: iLetter = `${quydoi[miniGroup[1]]} mươi ${iLetter}`; break;
        }
        switch(miniGroup[0]){
            case "0":
                if (index!==0 && (miniGroup[1]!=="0" || miniGroup[2]!=="0")){
                    iLetter = `không trăm ${iLetter}`
                };
            break;
            default: iLetter = `${quydoi[miniGroup[0]]} trăm ${iLetter}`; break;
        }
        return iLetter;
    })
    const donvi = ["", "nghìn", "triệu", "tỷ"];
    let ketqua = "";
    let dem = 0;
    for (let i=groupChu.length; i!==0; i--){
        if(groupChu[i-1].trim()!==""){
            ketqua= `${groupChu[i-1]} ${donvi[dem]} ${ketqua.trim()}`;
        };
        if (dem <3){dem++}else{dem = 1};
    }
    if (ketqua === ""){ketqua= "không"}
    return(ketqua);
}

