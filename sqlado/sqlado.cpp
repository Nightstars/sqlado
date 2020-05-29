// sqlado.cpp : 此文件包含 "main" 函数。程序执行将在此处开始并结束。
//

#include "iostream"  
#include "string"  
#include "vector"  
//步骤1：添加对ADO的支持  
#import "C:/Program Files/Common Files/System/ado/msado15.dll" no_namespace rename("EOF","adoEOF")rename("BOF","doBOF")

using namespace std;

int main()
{
	CoInitialize(NULL); //初始化COM环境           
	_ConnectionPtr pMyConnect(__uuidof(Connection));//定义连接对象并实例化对象 
	_RecordsetPtr pRst(__uuidof(Recordset));//定义记录集对象并实例化对象               
	try
	{
		//步骤2：创建数据源连接
		/*打开数据库“SQLServer”，这里需要根据自己PC的数据库的情况 */
		pMyConnect->Open("Provider=SQLOLEDB; Server=192.168.0.187,1433\MSSQLSERVER;Database=CMS; uid=sa; pwd=Ihavenoidea@0;", "", "", adModeUnknown);
		//注意：计算机全名可以在计算机的属性查看，SchoolTemp是数据库名，账户如sa
	}
	catch (_com_error& e)
	{
		cout << "Initiate failed!" << endl;
		cout << e.Description() << endl;
		cout << e.HelpFile() << endl;
		return 0;
	}
	cout << "Connect succeed!" << endl;

	//步骤3：对数据源中的数据库/表进行操作
	try
	{
		pMyConnect->Execute("use CMS", NULL, adCmdText);//执行SQL： select * from gendat 
		pRst = pMyConnect->Execute("select top(100) * from cop_act_head", NULL, adCmdText);//执行SQL： select * from gendat 
		//Table_1是数据库SchoolTemp的表名
		if (!pRst->adoEOF)
		{
			pRst->MoveFirst();
		}
		else
		{
			cout << "Data is empty!" << endl;
			return 0;
		}
		vector<_bstr_t> column_name;

		//	/*存储表的所有列名，显示表的列名*/
		for (int i = 0; i < pRst->Fields->GetCount(); i++)
		{
			cout << pRst->Fields->GetItem(_variant_t((long)i))->Name << endl;
			column_name.push_back(pRst->Fields->GetItem(_variant_t((long)i))->Name);
		}
		cout << endl;

		//	/*对表进行遍历访问,显示表中每一行的内容*/
		while (!pRst->adoEOF)
		{
			vector<_bstr_t>::iterator iter = column_name.begin();
			for (iter; iter != column_name.end(); iter++)
			{
				if (pRst->GetCollect(*iter).vt != VT_NULL)
				{
					cout << (_bstr_t)pRst->GetCollect(*iter) << endl;
				}
				else
				{
					cout << "NULL" << endl;
				}
			}
			pRst->MoveNext();
			cout << endl;
		}
	}
	catch (_com_error& e)
	{
		cout << e.Description() << endl;
		cout << e.HelpFile() << endl;
		return 0;
	}

	//步骤4：关闭数据源
	/*关闭数据库并释放指针*/
	try
	{
		pRst->Close();     //关闭记录集               
		pMyConnect->Close();//关闭数据库               
		pRst.Release();//释放记录集对象指针               
		pMyConnect.Release();//释放连接对象指针
	}
	catch (_com_error& e)
	{
		cout << e.Description() << endl;
		cout << e.HelpFile() << endl;
		return 0;
	}
	CoUninitialize(); //释放COM环境
	system("pause");
	return 0;
}

// 运行程序: Ctrl + F5 或调试 >“开始执行(不调试)”菜单
// 调试程序: F5 或调试 >“开始调试”菜单

// 入门使用技巧: 
//   1. 使用解决方案资源管理器窗口添加/管理文件
//   2. 使用团队资源管理器窗口连接到源代码管理
//   3. 使用输出窗口查看生成输出和其他消息
//   4. 使用错误列表窗口查看错误
//   5. 转到“项目”>“添加新项”以创建新的代码文件，或转到“项目”>“添加现有项”以将现有代码文件添加到项目
//   6. 将来，若要再次打开此项目，请转到“文件”>“打开”>“项目”并选择 .sln 文件
