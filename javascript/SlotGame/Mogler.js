
/*export default*/ class Mogler {

    //----------メンバ変数----------//
    //-----クラス変数
    //合計フラグ数（クジの枚数）
    static #FLAG_SUM = 65535;
    //固定のフラグ（確率はFLAG_SUM÷COMMON_FLAG[key]）
    static #COMMON_FLAG = {Rep:8970, Bell:60, Pierrot:60, CherryBig:30, CherryReg:30, CenterCherry:6, Freeze:1};
    //設定毎のフラグ数
    static #LEVEL_FLAG = {
        //[Grape, Cherry, Big, Reg]
        1:{Grape:9320, Cherry:1750, Big:200, Reg:120},
        2:{Grape:9400, Cherry:1760, Big:200, Reg:140},
        3:{Grape:9450, Cherry:1850, Big:210, Reg:150},
        4:{Grape:9500, Cherry:1900, Big:220,  Reg:180},
        5:{Grape:9600, Cherry:1900, Big:230,  Reg:210},
        6:{Grape:9800, Cherry:1950, Big:230,  Reg:240}
    };
    //各フラグのアウト枚数、文字列はプレミアフラグの識別子
    static #FLAG_VALUE = {
        Grape:7, Cherry:2, Big:312, Reg:104,
        Rep:3, Bell:15, Pierrot:15, CherryBig:312, 
        CherryReg:104, CenterCherry:7, Freeze:7, noLot:0
    };
    //継続率（%）
    static #STOCK_PROB = 75;
    //対応するストックフラグ
    static #STOCK_MODE = {
            CenterCherry: "Reg",
            Freeze: "Big"
    };

    //-----インスタンス変数
    #flagItem = {};                 //インスタンスに設定されたフラグ
    #level = 0;                     //台の設定
    #outRe = 0;                     //現在、若しくは直近のメダルアウト
    #stockCnt = {Big:0, Reg:0};      //ストックしているボーナスフラグ

    //台の持つ履歴
    #totalSpin = 0;
    #nowSpin = 0;
    #reSpin = [];
    #flagCnt = {Grape:0, Cherry:0, Big:0, Reg:0,
        Rep:0, Bell:0, Pierrot:0, CherryBig:0, 
        CherryReg:0, CenterCherry:0, Freeze:0, noLot:0
    };
    

    //----------コンストラクター----------//
    //arg1 : インスタンス化する台の設定
    constructor(level){
        //設定 1～6のみ
        if (!Number.isInteger(level) || level < 1 || level > 6) {
            throw new Error("Invalid input: level must be 1–6.");
        } else {
            this.#level = level;
            //flagItemに設定別でフラグ抽選確率をセットする
            const res = Object.assign({}, Mogler.#COMMON_FLAG, Mogler.#LEVEL_FLAG[this.#level]);
            Object.entries(res).reduce((acc, [key, val]) => {
                acc += val;//累積
                this.#flagItem[key] = acc;
                return acc;
            }, 0);
        }
    }

    //----------パブリックメソッド----------//
    //-----回転と結果
    //-----arg無し
    spinResult() {

        //回転結果
        const flgRes = Math.floor(Math.random() * Mogler.#FLAG_SUM);
        
        //フラグ格納変数の宣言と共に、インスタンスにストックがある場合格納
        let key = ["Big", "Reg"].find(stockFlg => this.#stockCnt[stockFlg] >= 1);

        //ストックがある場合、連想配列先頭から順にストックを消費して抽選をスキップ
        if (key) {
            this.#stockCnt[key]--;
        } else {
            //回転結果からインスタンスが持つフラグと照合
            key = Object.entries(this.#flagItem)
                .find(([_, val]) => flgRes <= val)?.[0] ?? "noLot";
        }

        //プレミアフラグ判定
        if (key in Mogler.#STOCK_MODE) {
            this.#stockGame(Mogler.#STOCK_MODE[key]);
        }
        this.#outRe = Mogler.#FLAG_VALUE[key];
        //データ履歴を付ける
        this.#dataCnt(key);

        return {flag: key, payout: this.#outRe};
    }

    //-----台の持つ履歴へのアクセサ
    //-----arg1 : 文字列によって取得するデータを選択できる
    getData(dataName) {
        switch(dataName) {
            case "total":   return this.#totalSpin;
            case "spin":    return this.#nowSpin;
            case "history": return this.#reSpin;
            case "flg":     return this.#flagCnt;          
            case "stock":   return Object.values(this.#stockCnt).some(cnt => cnt >= 1); //ストックが1以上あるならtrueを返却
            default:
                throw new Error(`Unknown dataName: ${dataName}`);
        }
    }

    //----------プライベートメソッド----------//
    //-----ストックモード
    //-----arg1 : Big連かReg連かを設定
    #stockGame(flg) {

        //ループストックをメンバ変数に代入
        const loopStock = (total = 0) => 
            Math.random() * 100 < Mogler.#STOCK_PROB ? loopStock(++total) : total;
        this.#stockCnt[flg] = Number(loopStock());
    }

    //-----データカウンター
    //-----arg1 : データに追加するフラグ
    #dataCnt(flg) {
        this.#totalSpin++;        //総回転数
        this.#nowSpin++;          //現在回転数
        this.#flagCnt[flg]++;     //フラグ回数を追加

        //当選フラグを引いた場合、現在の回転数をリセットし、履歴の配列に当選時の回転数を記録
        switch(flg) {
            case "Big":
            case "Reg":
            case "CenterCherry":
            case "Freeze":
                //台の当選履歴を直近10回までとし、現在回転数を初期化
                this.#reSpin.push(this.#nowSpin);
                this.#reSpin.length > 10 &&  this.#reSpin.shift();
                this.#nowSpin = 0;
                break;
            default:
                //pass
        }
    } 
}
