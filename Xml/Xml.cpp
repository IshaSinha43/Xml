
#include <iostream>
#include <iostream>
#include <atlbase.h>
#include <comutil.h>
#import <msxml6.dll>

using namespace MSXML2;
using namespace std;


HRESULT paramInput(MSXML2::IXMLDOMDocument2Ptr& docPtr,
	MSXML2::IXMLDOMElementPtr& paramElement,
	CComBSTR bstrValue1,
	CComBSTR bstrValue2,
	CComBSTR bstrValue3,
	CComBSTR bstrValue4,
	CComBSTR bstrValue5,
	CComVariant varValue1,
	CComVariant varValue2,
	CComVariant varValue3,
	CComVariant varValue4)
{
	try
	{

		MSXML2::IXMLDOMElementPtr paramElements = docPtr->createElement(L"param");
		paramElements->raw_setAttribute(bstrValue1.m_str, varValue1);
		paramElements->raw_setAttribute(bstrValue2.m_str, varValue2);
		paramElements->raw_setAttribute(bstrValue3.m_str, varValue3);
		paramElements->raw_setAttribute(bstrValue4.m_str, varValue4);
		paramElements->put_text(bstrValue5.m_str);
		paramElement->appendChild(paramElements);
	}
	catch (...)
	{
	}
	return S_OK;
}
HRESULT CreateElement(MSXML2::IXMLDOMDocument2Ptr& docPtr,
	MSXML2::IXMLDOMElementPtr& parentElement,
	MSXML2::IXMLDOMElementPtr& newElement,
	CComBSTR bstrElementName,
	CComBSTR bstrText)
{
	try
	{
		if (bstrElementName.Length())
		{
			newElement = docPtr->createElement(bstrElementName.m_str);
			if (newElement)
			{
				if (bstrText.Length())
					newElement->put_text(bstrText.m_str);

				parentElement->appendChild(newElement);
			}
		}

	}
	catch (...)
	{
	}
	return S_OK;
}



HRESULT setAttribute(

	MSXML2::IXMLDOMElementPtr& objectElement,
	CComBSTR bstrclsid,
	CComBSTR bstrname,
	CComBSTR bstrtype,
	CComBSTR bstrid,
	CComVariant varclsid,
	CComVariant varname,
	CComVariant vartype,
	CComVariant varid)
{
	try
	{
		objectElement->raw_setAttribute(bstrclsid.m_str, varname);
		objectElement->raw_setAttribute(bstrname.m_str, varname);
		objectElement->raw_setAttribute(bstrtype.m_str, vartype);
		objectElement->raw_setAttribute(bstrid.m_str, varid);
	}
	catch (...)
	{
	}
	return S_OK;
}

int WriteXml()
{
	try
	{
		//Initializing COM 
		::CoInitialize(NULL);


		//Creating dom document 
		MSXML2::IXMLDOMDocument2Ptr docPtr;
		HRESULT hr = docPtr.CreateInstance(__uuidof(MSXML2::DOMDocument60));
		if (hr != S_OK || !docPtr)
		{
			return 0;
		}

		//Loading dom document
		CComBSTR bstrXML(L"<objects></objects>");
		CComVariant varValue;
		CComBSTR bstrValue;
		VARIANT_BOOL vbRet = VARIANT_FALSE;
		docPtr->raw_loadXML(bstrXML.m_str, &vbRet);

		if (vbRet == VARIANT_TRUE)
		{
			// MSXML2::IXMLDOMElementPtr objectElement;

			// CreateElement(docPtr, (MSXML2::IXMLDOMElementPtr&)docPtr->documentElement, objectElement, L"id", L"object", L"name", L"clsid", L"type", L"" ,L"{46AEFCE4-FBE7-4DD6-8C4C-7653CB988C2C}", L"OBJECT", L"Background", L"Sceneobject");


			//Creating List
			MSXML2::IXMLDOMElementPtr objectElement = docPtr->createElement(L"object");

			docPtr->documentElement->appendChild(objectElement);

			objectElement->setAttribute(L"clsid", L"SceneObject");
			objectElement->setAttribute(L"name", L"Background");
			objectElement->setAttribute(L"type", L"object");
			objectElement->setAttribute(L"id", L"{46AEFCE4-FBE7-4DD6-8C4C-7653CB988C2C}");

			// MSXML2::IXMLDOMElementPtr ObjectElement = docPtr->createElement(L"objectst");
			 //objectInput(docPtr,ObjectElement,L"name",L"id", L"type",L"SceneObject");

			MSXML2::IXMLDOMElementPtr paramElement = docPtr->createElement(L"params");
			objectElement->appendChild(paramElement);

			paramInput(docPtr, paramElement, L"numkey", L"interpolator", L"type", L"name", L"Background", L"-1", L"NULL", L"STRING", L"name");

			paramInput(docPtr, paramElement, L"numkey", L"interpolator", L"type", L"name", L"", L"-1", L"NULL", L"texture", L"texturenodeid");

			docPtr->save(L"D:\\test.xml");
		}

		else
		{
			cout << "Fail to create xml.";
		}
	}
	catch (...)
	{
	}
	return 0;

}

int ParseXml()
{
	try
	{
		// Initializing COM 
		::CoInitialize(NULL);

		//Creating dom document 
		MSXML2::IXMLDOMDocument2Ptr docPtr;
		MSXML2::IXMLDOMElementPtr paramElement;
		MSXML2::IXMLDOMAttributePtr paramAttribute;
		HRESULT hr = docPtr.CreateInstance(__uuidof(MSXML2::DOMDocument60));
		if (hr != S_OK || !docPtr)
		{
			return 0;
		}

		//Loading dom document


		CComVariant varValue;
		CComBSTR bstrValue;

		VARIANT_BOOL vbRet = VARIANT_FALSE;
		hr = docPtr->raw_load(CComVariant(L"d:\\sampletest.xml"), &vbRet);

		if (hr == S_OK && vbRet == VARIANT_TRUE)//VARIANT_TRUE=-1
		{
			MSXML2::IXMLDOMNodePtr nodeObj;
			MSXML2::IXMLDOMElementPtr elementObj;
			MSXML2::IXMLDOMNodeListPtr nodeListParam;

			docPtr->documentElement->raw_selectSingleNode(CComBSTR(L"object").m_str, &nodeObj);

			elementObj = nodeObj;
			nodeListParam = elementObj->selectNodes(L"params/param");
			if (hr == S_OK && nodeListParam)
			{
				//Iterating list
				for (int cntr = 0; cntr < nodeListParam->length; cntr++)
				{
					MSXML2::IXMLDOMElementPtr elementParam = nodeListParam->item[cntr];
					if (elementParam)
					{
						hr = elementParam->raw_getAttribute(CComBSTR("type"), &varValue);
						if (hr == S_OK && varValue.bstrVal)
							bstrValue = varValue.bstrVal;

						bstrValue.Empty();
						bstrValue = elementParam->Gettext().operator wchar_t* ();

						hr = elementParam->raw_getAttribute(CComBSTR(L"name"), &varValue);
						if (hr == S_OK && varValue.bstrVal)
						{
							bstrValue = varValue.bstrVal;
						}
						hr = elementParam->raw_getAttribute(CComBSTR(L"interpolator"), &varValue);
						if (hr == S_OK && varValue.bstrVal)
						{
							bstrValue = varValue.bstrVal;
						}
						hr = elementParam->raw_getAttribute(CComBSTR(L"numkey"), &varValue);
						if (hr == S_OK && varValue.bstrVal)
						{
							bstrValue = varValue.bstrVal;
						}
					}

				}
			}
			else
			{
				cout << "Failed to load xml.";
			}

		}
	}
	catch (...)
	{
	}
	return 0;

}
int main()
{
	ParseXml();
	//	WriteXml();
	return 0;
}
