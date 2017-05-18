/*---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
 *  Author      : Andy Lin
 *
 *  Action      : Create at 2010/06/14
 *  Description : First version of development
 *
 *  Action      : Modified at 2010/06/18
 *  Description : 修正 StringValidationAfterTrim 的控制模式
 *---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
 *【Prototype Functions】
 *-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
String.prototype.Trim = function()
{
    return this.replace(/^\s+|\s+$/g, '');
}
/*---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
 *【JSON 結構內容說明】
 *---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
 *  - controls (Root)：
 *    封裝欲進行非空白輸入或數字字串內容驗證之控制項的相關資訊的容器。
 *
 *    - client_id (1st Layer)：
 *      欲被進行相關驗證輸入之控制項的 Client 端 ID。
 *
 *    - blank_validation (1st Layer)：
 *      (1) 用以定義是否進行非空白輸入的驗證。(此項目可為 Optional 選項)
 *      (2) 若欲使用此項目，請繼續遵循下列階層設定；且此集合中，僅能有唯一的一組設定項目。
 *
 *      - required (2nd Layer)：
 *        (1) 是否進行非空白輸入的驗證。(值：true/false)
 *        (2) 若此項元素不存在，則視為不進行非空白輸入驗證。
 *
 *      - message (2nd Layer)：
 *        (1) 未通過非空白輸入驗證時，所希望呈現的錯誤訊息。
 *        (2) 若此項元素不存在，則將呈現系統預設錯誤訊息。(值：[SYS: Blank Default Error Message !])
 *
 *    - numeric_validation (1st Layer)：
 *      (1) 用以定義是否進行數字字串格式的驗證。(此項目可為 Optional 選項)
 *      (2) 若欲使用此項目，請繼續遵循下列階層設定；且此集合中，僅能有唯一的一組設定項目。
 *
 *      - required (2nd Layer)：
 *        (1) 是否進行數字字串的驗證。(值：true/false)
 *        (2) 若此項元素不存在，則視為不進行數字字串驗證。
 *
 *      - message (2nd Layer)：
 *        (1) 當數字字串驗證未通過時，所希望呈現的錯誤訊息。
 *        (2) 若此項元素不存在，則將呈現系統預設錯誤訊息。(值：[SYS: Numeric Default Error Message !])
 *-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
function StringValidationAfterTrim(json_element)
{
    var result = '';

    if (json_element && json_element.controls && (json_element.controls.length > 0))
    {
        var control = null;
        var element = null;

        for (var i = 0; i < json_element.controls.length; ++i)
        {
            control = json_element.controls[i];

            if (control.client_id && (control.client_id.Trim() != ''))
            {
                element = document.getElementById(control.client_id);

                if (element)
                {
                    // 進行非空白輸入驗證
                    if (control.blank_validation)
                    {
                        if (control.blank_validation[0].required || (control.blank_validation[0].required == undefined))
                        {
                            if (element.value.Trim() == '')
                            {
                                if (control.blank_validation[0].message)
                                    result += '- ' + control.blank_validation[0].message + ' !\n';
                                else
                                    result += '- SYS: Blank {ER}\n';
                            }
                        }
                    }

                    // 進行非法字元輸入驗證
                    if (control.RequestFormError_validation) {
                        var re = /</;
                        if (control.RequestFormError_validation[0].required || (control.RequestFormError_validation[0].required == undefined)) {
                            if (re.test(element.value.Trim())) {
                                if (control.RequestFormError_validation[0].message)
                                    result += '- ' + control.RequestFormError_validation[0].message + ' !\n';
                                else
                                    result += '- SYS: Blank {ER}\n';
                            }
                        }
                    }

                    // 進行數字字串格式驗證
                    if (control.numeric_validation)
                    {
                        var NumericValidation = function()
                                                {
                                                    if (element.value.Trim() != '')
                                                    {
                                                        if (control.numeric_validation[0].required || (control.numeric_validation[0].required == undefined))
                                                        {
                                                            var is_numeric = /^[0-9]{1,}(\.[0-9]+)?$/.test(element.value.Trim());

                                                            if(!is_numeric)
                                                            {
                                                                if (control.numeric_validation[0].message)
                                                                    return '- ' + control.numeric_validation[0].message + ' !\n';
                                                                else
                                                                    return '- SYS: Numeric {ER}\n';
                                                            }
                                                        }
                                                    }

                                                    return '';
                                                };

                        result += NumericValidation();
                    }
                }
            }
        }

        control = null;
        element = null;
    }

    if (result)
    {
        alert(result.replace(/\{ER\}/g, 'Default Error Message !'));
        return false;
    }

    return true;
}
//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------