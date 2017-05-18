/*---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
 *  Author      : Andy Lin
 *
 *  Action      : Create at 2010/06/14
 *  Description : First version of development
 *
 *  Action      : Modified at 2010/06/18
 *  Description : �ץ� StringValidationAfterTrim ������Ҧ�
 *---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
 *�iPrototype Functions�j
 *-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
String.prototype.Trim = function()
{
    return this.replace(/^\s+|\s+$/g, '');
}
/*---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
 *�iJSON ���c���e�����j
 *---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
 *  - controls (Root)�G
 *    �ʸ˱��i��D�ťտ�J�μƦr�r�ꤺ�e���Ҥ������������T���e���C
 *
 *    - client_id (1st Layer)�G
 *      ���Q�i��������ҿ�J������� Client �� ID�C
 *
 *    - blank_validation (1st Layer)�G
 *      (1) �ΥH�w�q�O�_�i��D�ťտ�J�����ҡC(�����إi�� Optional �ﶵ)
 *      (2) �Y���ϥΦ����ءA���~���`�U�C���h�]�w�F�B�����X���A�ȯ঳�ߤ@���@�ճ]�w���ءC
 *
 *      - required (2nd Layer)�G
 *        (1) �O�_�i��D�ťտ�J�����ҡC(�ȡGtrue/false)
 *        (2) �Y�����������s�b�A�h�������i��D�ťտ�J���ҡC
 *
 *      - message (2nd Layer)�G
 *        (1) ���q�L�D�ťտ�J���ҮɡA�ҧƱ�e�{�����~�T���C
 *        (2) �Y�����������s�b�A�h�N�e�{�t�ιw�]���~�T���C(�ȡG[SYS: Blank Default Error Message !])
 *
 *    - numeric_validation (1st Layer)�G
 *      (1) �ΥH�w�q�O�_�i��Ʀr�r��榡�����ҡC(�����إi�� Optional �ﶵ)
 *      (2) �Y���ϥΦ����ءA���~���`�U�C���h�]�w�F�B�����X���A�ȯ঳�ߤ@���@�ճ]�w���ءC
 *
 *      - required (2nd Layer)�G
 *        (1) �O�_�i��Ʀr�r�ꪺ���ҡC(�ȡGtrue/false)
 *        (2) �Y�����������s�b�A�h�������i��Ʀr�r�����ҡC
 *
 *      - message (2nd Layer)�G
 *        (1) ��Ʀr�r�����ҥ��q�L�ɡA�ҧƱ�e�{�����~�T���C
 *        (2) �Y�����������s�b�A�h�N�e�{�t�ιw�]���~�T���C(�ȡG[SYS: Numeric Default Error Message !])
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
                    // �i��D�ťտ�J����
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

                    // �i��D�k�r����J����
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

                    // �i��Ʀr�r��榡����
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