//<ref>Newtonsoft.Json.dll</ref>

using System;

using Mongoose.IDO.Protocol;
using Mongoose.Scripting;
using Mongoose.Core.Common;
using Mongoose.Core.DataAccess;

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;

using Newtonsoft.Json;

namespace Mongoose.FormScripts
{
   public class ItemPicking : FormScript
   {
      /// <summary>
      /// グローバル変数
      /// </summary>
      string gIDOName = "SLLotLocs";  // ターゲットIDO名
      string gSSO = "1";  // SSOの使用、基本は1　1：はい　0：いいえ
      string gServerId = "0";  // APIサーバのID、詳しくはAPIサーバ関連
      string gSuiteContext = "CSI";  // APIスイートのコンテキスト、同上
      string gContentType = "application/json";  // 返り値のタイプ、基本設定不要もしくは"application/json"
      string gTimeout = "10000";  // タイムアウト設定
      string gSiteName = "Q72Q74BY8XUT3SKY_TRN_AJP";  // ターゲットIDOが存在するサイト名


      /// <summary>
      /// 取得データ、更新データ作成のクラス
      /// </summary>
      public class cItem
      {
         // 取得予定のプロパティ
         [JsonProperty("Whse")]
         public string Whse { get; set; }

         [JsonProperty("Lot")]
         public string Lot { get; set; }

         [JsonProperty("Loc")]
         public string Loc { get; set; }

         [JsonProperty("QtyOnHand")]
         public string QtyOnHand { get; set; }

         [JsonProperty("PickingNum")]
         public string PickingNum { get; set; }

         [JsonProperty("Item")]
         public string Item { get; set; }

         [JsonProperty("ItemDescription")]
         public string ItemDescription { get; set; }

         // 更新時必要のプロパティ（基本IDO作成時自動生成される）
         [JsonProperty("RecordDate")]
         public string RecordDate { get; set; }

         [JsonProperty("RowPointer")]
         public string RowPointer { get; set; }

         [JsonProperty("_ItemId")]
         public string _ItemId { get; set; }

      }

      /// <summary>
      /// APIを呼び出してデータ取得・グリッド表示を行う
      /// </summary>
      /// <param name="mode">操作モード（0:取得, 1:挿入, 2:更新, 4:削除）</param>
      public void callAPI(int mode){
         callAPIwithTarget(mode,"");
      }
      
      /// <summary>
      /// APIを呼び出してデータ取得・グリッド表示を行う
      /// </summary>
      /// <param name="mode">操作モード（0:取得, 1:挿入, 2:更新, 4:削除）</param>
      /// <param name="target">更新目標（今はLot）</param>
      public void callAPIwithTarget(int mode,string target){

         List<Property> propertiesList = new List<Property>();

         //使用するプロパティのデータリストを作成
         if(mode == 0){
            propertiesList.Add(new Property { Name = "Whse",Value = "", Modified = true });
            propertiesList.Add(new Property { Name = "Lot",Value = "", Modified = true });
            propertiesList.Add(new Property { Name = "Loc",Value = "", Modified = true });
            propertiesList.Add(new Property { Name = "QtyOnHand",Value = "", Modified = true });
            propertiesList.Add(new Property { Name = "Item",Value = "", Modified = true });
            propertiesList.Add(new Property { Name = "ItemDescription",Value = "", Modified = true });
            propertiesList.Add(new Property { Name = "RecordDate",Value = "", Modified = true });
            propertiesList.Add(new Property { Name = "RowPointer",Value = "", Modified = true });
            propertiesList.Add(new Property { Name = "_ItemId",Value = "", Modified = true });

         }
         //データ更新用、get時不要
         else{
            propertiesList.Add(new Property { Name = "Whse", Value = ThisForm.Components["ResultGrid"].GetGridValue(ThisForm.Components["ResultGrid"].GetGridCurrentRow(),1), Modified = true});
            propertiesList.Add(new Property { Name = "Lot", Value = ThisForm.Components["ResultGrid"].GetGridValue(ThisForm.Components["ResultGrid"].GetGridCurrentRow(),2), Modified = true});
            propertiesList.Add(new Property { Name = "Loc", Value = ThisForm.Components["ResultGrid"].GetGridValue(ThisForm.Components["ResultGrid"].GetGridCurrentRow(),3), Modified = true});
            propertiesList.Add(new Property { Name = "QtyOnHand", Value = ThisForm.Components["ResultGrid"].GetGridValue(ThisForm.Components["ResultGrid"].GetGridCurrentRow(),4), Modified = true});
            propertiesList.Add(new Property { Name = "Item", Value = ThisForm.Components["ResultGrid"].GetGridValue(ThisForm.Components["ResultGrid"].GetGridCurrentRow(),5), Modified = true});
            propertiesList.Add(new Property { Name = "ItemDescription", Value = ThisForm.Variables("gItemDescription").Value, Modified = true});
            //データ特定用、insert時は不要
            if(mode != 1){
               propertiesList.Add(new Property { Name = "RecordDate", Value = ThisForm.Components["ResultGrid"].GetGridValue(ThisForm.Components["ResultGrid"].GetGridCurrentRow(),9), Modified = true });
               propertiesList.Add(new Property { Name = "RowPointer", Value = ThisForm.Components["ResultGrid"].GetGridValue(ThisForm.Components["ResultGrid"].GetGridCurrentRow(),10), Modified = true });
               propertiesList.Add(new Property { Name = "_ItemId", Value = ThisForm.Components["ResultGrid"].GetGridValue(ThisForm.Components["ResultGrid"].GetGridCurrentRow(),11), Modified = true });
            }
         }

         //データを取得
         string JSONResponse = getData(mode,gIDOName,propertiesList,target);
         
         //取得した結果はJSONなので、bodyを処理しデータを取得
         List<cItem> itemsList = JsonConvert.DeserializeObject<getJSONObject>(JSONResponse).Items;

         //処理結果を変数に格納
         ThisForm.Variables("vJSONResult").Value = GenerateWebSetJson(itemsList);

         //グリッドの初期化
         int count = 1;
         if(ThisForm.Components["ResultGrid"].GetGridRowCount() > 0 && target == ""){
            ThisForm.Components["ResultGrid"].DeleteGridRows(1,ThisForm.Components["ResultGrid"].GetGridRowCount());
         }

         //データを表示
         foreach (var item in itemsList)
         {
            ThisForm.Components["ResultGrid"].InsertGridRows(count,1);
            ThisForm.Components["ResultGrid"].SetGridValue(count,1,item.Whse);
            ThisForm.Components["ResultGrid"].SetGridValue(count,2,item.Lot);
            ThisForm.Components["ResultGrid"].SetGridValue(count,3,item.Loc);
            ThisForm.Components["ResultGrid"].SetGridValue(count,4,float.Parse(item.QtyOnHand).ToString());//小数点消すためいったんfloatに変換
            //ThisForm.Components["ResultGrid"].SetGridValue(count,5,item.PickingNum);
            ThisForm.Components["ResultGrid"].SetGridValue(count,6,item.Item);
            ThisForm.Components["ResultGrid"].SetGridValue(count,7,item.ItemDescription);
            //ThisForm.Components["ResultGrid"].SetGridValue(count,8,item.Match);
            ThisForm.Components["ResultGrid"].SetGridValue(count,9,item.RecordDate);
            ThisForm.Components["ResultGrid"].SetGridValue(count,10,item.RowPointer);
            ThisForm.Components["ResultGrid"].SetGridValue(count,11,item._ItemId);
            
            count++;
         } 
      }


      /// <summary>
      /// データをアップデート・取得する
      /// </summary>
      /// <param name="mode">操作モード（0:取得, 1:挿入, 2:更新, 4:削除）</param>
      /// <param name="IDOName">IDO名</param>
      /// <param name="propertiesList">プロパティリスト</param>
      /// <param name="target">更新目標（今はLot）</param>
      /// <returns>APIレスポンスのJSON文字列</returns>
      private string getData(int mode ,string IDOName ,List<Property> propertiesList,string target){
         // 呼び出すAPIを設定
         string sso = gSSO;                        // SSOの使用、基本は1　1：はい　0：いいえ
         string serverId = gServerId;              // APIサーバのID、詳しくはAPIサーバ関連
         string suiteContext = gSuiteContext;      // APIスイートのコンテキスト、同上
         string httpMethod = "";                   // APIメソッドのHTTP接続し設定、同上
         string methodName = "";                   // APIメソッドの名前、同上、IDO名以降の部分は動的にできそう
         string parameters = "";                   // APIパラメータ設定、一部パラメータはメソッド名で設定できるが、ここはURL上で設定できない部分を書く（ヘッダーなど）
         string contentType = gContentType;        // 返り値のタイプ、基本設定不要もしくは"application/json"
         string timeout = gTimeout;                // タイムアウト設定

         // 使用するメソッド（APIのエンドポイント）を設定
         if(mode == 0){
            // getの場合
            httpMethod = "GET";
            // APIのURL + IDO名 + 取得プロパティ
            if(target == ""){
               // 通常取得
               methodName = $"/IDORequestService/ido/load/{IDOName}?properties={string.Join( ", ", propertiesList.Select(p => p.Name) )}{GenerateFilter()}";
            }
            else{
               // 特定取得
               methodName = $"/IDORequestService/ido/load/{IDOName}?properties={string.Join( ", ", propertiesList.Select(p => p.Name) )}&filter=Lot=\'{target}\'";
            }
         }
         else{
            // postの場合
            httpMethod = "POST";
            // APIのURL + IDO名 + 取得パラメータ
            methodName = $"/IDORequestService/ido/update/{IDOName}?refresh=true";
         }

         // パラメータのHeader部分（基本共通）
         string headerTemp = $"{{\"Name\":\"X-Infor-MongooseConfig\",\"Type\":\"Header\",\"Value\":\"{gSiteName}\"}}";

         // パラメータを設定
         if(mode == 0){
            // getの場合
            // 基本Headerのみ
            parameters = $"[{headerTemp}]";
         }
         else{
            // postの場合
            // Body部分、動的生成する必要がある
            string bodyTemp = $"{{\"Name\":\"Body\",\"Type\":\"Body\",\"Value\":\'{GenerateChangeSetJson(mode,IDOName,propertiesList)}\'}}";

            // 結合
            parameters = $"[{headerTemp},{bodyTemp}]";;
            
         }

         // API呼び出し
         InvokeRequestData IDORequest = new InvokeRequestData();  //APIメソッド呼び出し用
         InvokeResponseData response;                             //結果受け取り用

         // API呼び出すためのIDOメソッドを設定
         // APIを直で呼び出せないため、IDOメソッドを経由
         // また、LoadCollectionメソッドは同アプリ（mongoose）同サイト（default）のみ呼び出すため不採用
         IDORequest.IDOName = "IONAPIMethods";
         IDORequest.MethodName = "InvokeIONAPIMethod";

         // パラメータ設定
         IDORequest.Parameters.Add(new InvokeParameter(sso));
         IDORequest.Parameters.Add(new InvokeParameter(serverId));
         IDORequest.Parameters.Add(new InvokeParameter(suiteContext));
         IDORequest.Parameters.Add(new InvokeParameter(httpMethod));
         IDORequest.Parameters.Add(new InvokeParameter(methodName));
         IDORequest.Parameters.Add(new InvokeParameter(parameters));
         IDORequest.Parameters.Add(new InvokeParameter(contentType));
         IDORequest.Parameters.Add(new InvokeParameter(timeout));
         IDORequest.Parameters.Add(new InvokeParameter(IDONull.Value)); //ResponseCode
         IDORequest.Parameters.Add(new InvokeParameter(IDONull.Value)); //ResponseContent
         IDORequest.Parameters.Add(new InvokeParameter(IDONull.Value)); //ResponseHeaders
         IDORequest.Parameters.Add(new InvokeParameter(IDONull.Value)); //ResponseInfobar

         // メソッド呼び出し
         response = IDOClient.Invoke(IDORequest);

         // エラーハンドラ
         if (response.IsReturnValueStdError())
         {
            // get infobar output parameter
            string errorMsg = null;
            errorMsg = response.Parameters[8].Value;
            errorMsg += "\r\n" + response.Parameters[11].Value;
            throw new Exception(string.Format("Trigger of IONAPI Method Failed : {0}", errorMsg));
         }

         // getの場合
         if(mode == 0){
            //データを画面に表示
            return response.Parameters[9].Value;
         }

         // postの場合
         // データや_itemIDの更新後の状況を取得するため、getで再帰呼び出し
         return getData(0, IDOName, propertiesList,"");
      }

      /// <summary>
      /// Update用JSON文字列を自動生成する
      /// </summary>
      /// <param name="mode">操作モード（0:取得, 1:挿入, 2:更新, 4:削除）</param>
      /// <param name="IDOName">IDO名</param>
      /// <param name="propertiesList">プロパティリスト</param>
      /// <returns>Update用JSON文字列</returns>
      private string GenerateChangeSetJson(int mode ,string IDOName ,List<Property> propertiesList)
      {
         // 1. ルートオブジェクトを作成
         var root = new UpdateJSONObject
         {
            IDOName = IDOName
         };

         // 2. Changeオブジェクトを作成
         var change = new Change
         {
            Action = mode,
            ItemId = "new1"
         };

         foreach(Property prop in propertiesList){
            // 3. Propertyオブジェクトを必要な分だけ作成し、Changeオブジェクトのリストに追加
            change.Properties.Add(prop);

            //キー指定する場合、ItemIDを指定されたキーと同値にする（mode = 2,4のみ）
            if(!prop.Modified && change.ItemId == "new1"){
               change.ItemId = prop.Value;
            }
         }

         // 4. Changeオブジェクトをルートオブジェクトのリストに追加
         root.Changes.Add(change);

         // 5. オブジェクトをJSON文字列にシリアライズ
         // Formatting.Indented を付けると整形された見やすいJSONが出力される
         string jsonString = JsonConvert.SerializeObject(root, Formatting.Indented);

         return jsonString;
      }

      /// <summary>
      /// フィルター文字列を自動生成する
      /// </summary>
      /// <returns>フィルター文字列</returns>  
      private string GenerateFilter(){
         string header = "&filter=";
         string filter = "";

         // 検索欄から変数を獲得 gShippingLocation         
         // 出荷場所
         if(ThisForm.Variables("gWhse").Value != ""){
            // 接続詞追加
            if(filter != ""){filter += " and ";}
            // 検索値追加
            filter += $"Whse=\'{ThisForm.Variables("gWhse").Value}\'";
         }

         // 品目ID
            // 接続詞追加
            if(filter != ""){filter += " and ";}
            // 検索値追加
            filter += $"QtyOnHand>0";
         

         // return "";
         // 作成したフィルターを返す
         if (filter == ""){
            return "";
         }
         else
         {
            return header + filter;
         }
      }

      /// <summary>
      /// QR読み取り対象をGridに追加
      /// </summary>
      public void addScanTarget(){
         callAPIwithTarget(0,ThisForm.Variables("HiddenValue").Value);
      }

      /// <summary>
      /// Grid内にある品目IDを照合
      /// </summary>
      public void matchGridItems(){
         int count = ThisForm.Components["ResultGrid"].GetGridRowCount();
         if(count == 0)return;
         if(ThisForm.Variables("HiddenValue").Value == "")ThisForm.Variables("HiddenValue").Value = ThisForm.Variables("gItem").Value;
         for(int i = 1;i < count+1 ; i++){
            if(ThisForm.Components["ResultGrid"].GetGridValue(i,6) == ThisForm.Variables("HiddenValue").Value){
               ThisForm.Components["ResultGrid"].SetGridValue(i,8,"OK");
            }
            else{
               ThisForm.Components["ResultGrid"].SetGridValue(i,8,"NG");
            }
         }
      }

      /// <summary>
      /// web用データをJSON文字列を生成
      /// </summary>
      /// <param name="resultList">整形結果リスト</param>
      /// <returns>web用JSON文字列</returns>
      private string GenerateWebSetJson(List<cItem> resultList){
         // オブジェクトを作成
         var responseObject = new WebJSONObject
         {
            itemName = ThisForm.Variables("gDescription").Value ,
            itemId = ThisForm.Variables("gItem").Value ,
            itemQuantity = ThisForm.Variables("gQtyOrderedConv").Value ,
            Items = resultList
         };

         // Formatting.Indented で整形して出力
         string jsonString = JsonConvert.SerializeObject(responseObject, Formatting.Indented);

         return jsonString;
      }

      //API取得JSONリスト化用クラス

      /// <summary>
      /// JSONからデータ取得用のクラス
      /// </summary>
      public class getJSONObject
      {
         [JsonProperty("Items")]
         public List<cItem> Items { get; set; }
      }

      /// <summary>
      /// 一番内側の "Properties" 配列の要素を表すクラス
      /// </summary>
      public class Property
      {
         [JsonProperty("Modified")]
         public bool Modified { get; set; }

         [JsonProperty("Value")]
         public string Value { get; set; }

         [JsonProperty("IsNull")]
         public bool IsNull { get; set; } = false;

         [JsonProperty("Name")]
         public string Name { get; set; }
      }

      //updateプロパティ作成用クラス

      /// <summary>
      /// "Changes" 配列の要素を表すクラス
      /// </summary>
      public class Change
      {
         [JsonProperty("Action")]
         public int Action { get; set; }

         [JsonProperty("ItemNo")]
         public int ItemNo { get; set; } = 0;

         [JsonProperty("UpdateLocking")]
         public int UpdateLocking { get; set; } = 0;

         [JsonProperty("Properties")]
         public List<Property> Properties { get; set; } = new List<Property>();

         [JsonProperty("ItemId")]
         public string ItemId { get; set; }
      }

      /// <summary>
      /// JSON全体のルートオブジェクトを表すクラス
      /// </summary>
      public class UpdateJSONObject
      {
         [JsonProperty("Changes")]
         public List<Change> Changes { get; set; } = new List<Change>();

         [JsonProperty("IDOName")]
         public string IDOName { get; set; }
      }

      /// <summary>
      /// web対応JSON整形用クラス
      /// </summary>
      public class WebJSONObject
      {
         // 品目名
         [JsonProperty("itemName")]
         public string itemName { get; set; }

         // 品目ID
         [JsonProperty("itemId")]
         public string itemId { get; set; }

         // 品目数
         [JsonProperty("itemQuantity")]
         public string itemQuantity { get; set; }

         [JsonProperty("Items")]
         public List<cItem> Items { get; set; }
      }
   }
   
}












