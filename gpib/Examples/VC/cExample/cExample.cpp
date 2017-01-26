// cExample.cpp : Defines the entry point for the console application.
//


//Purpose:  To illustrate the use of the automation layer in a C++ program
//without the use of smart COM pointers.  (In general I would advise using smart 
//pointers but this example will allow the use of the automation layer without them.
//Note that this is a C++ program and not C, that is we use class structures, however
//it is 'bare bones' in the sense we use COM calls directly and no smart pointers)
//

//standard include for a VC++ project
#include "stdafx.h"

//need for COM
#include "windows.h"
#include "comdef.h"


//This is the 'Automation' object (COM class).  It contains the IIOManger interface
//below
const CLSID CLSID_AGTIOManager = {0xAFF8D8E1,0xCE0D,0x11D3,{0x98,0xBB,0x00,0x10,0x83,0x01,0xCB,0x39}};

//The IIOManager allows the opening of a device using its address
//in the 'ConnectToInstrument' function.  The IO manager class above exports this
//COM interface.
const IID IID_IIOManager = {0xAFF8D8F8,0xCE0D,0x11D3,{0x98,0xBB,0x00,0x10,0x83,0x01,0xCB,0x39}};

//The IIOManger function 'ConnectToInstrument' will return the IIO interface
//this is an instance of the instrument. 
const IID IID_IIO = {0xAFF8D8EE,0xCE0D,0x11D3,{0x98,0xBB,0x00,0x10,0x83,0x01,0xCB,0x39}};

//The following are the definitions needed to program the instrument.  They
//can be placed in a header file for but are shown here to allow ease of
//perusal 

/************************************************************************************/
//This definition is to complete the IIO interface below.  You needn't use this
//interface directly to communicate with instruments
extern "C" {

interface IComponent : IDispatch {
        STDMETHOD(ComponentManufacturer)( BSTR* pVal);
        STDMETHOD(ComponentDescription)( BSTR* Desc);
        STDMETHOD(ComponentVersion)( BSTR* pVal);
        STDMETHOD(ComponentProgID)(BSTR* pVal);
        STDMETHOD(LogInterface)( VARIANT_BOOL pVal);
        STDMETHOD(LogInterface)( VARIANT_BOOL* pVal);
        STDMETHOD(InstanceName)( BSTR* pVal);
        STDMETHOD(InstanceName)( BSTR pVal);
    };


//This is your instrument.   Use the Enter and Output to communicate with the
//instrument.  Clear does a device clear, Query can also be used to combine
//enter & output for query commands.   These are the generally useful entry
//points for instruments such as the 34401 voltmeter.
interface IIO : IComponent {
        STDMETHOD(ComponentManufacturer)( BSTR* pVal);
        STDMETHOD(ComponentDescription)( BSTR* Desc);
        STDMETHOD(ComponentVersion)( BSTR* pVal);
        STDMETHOD(ComponentProgID)(BSTR* pVal);
        STDMETHOD(LogInterface)( VARIANT_BOOL pVal);
        STDMETHOD(LogInterface)( VARIANT_BOOL* pVal);
        STDMETHOD(InstanceName)( BSTR* pVal);
        STDMETHOD(InstanceName)( BSTR pVal);
        STDMETHOD(CanHandleConnectionName)(BSTR ConnectionName,VARIANT_BOOL* HandleIt);
        STDMETHOD(Clear)();
        STDMETHOD(Connect)(VARIANT ConnectionName);
        STDMETHOD( ConnectionName)(BSTR* pVal);
        STDMETHOD(DeviceLock)();
        STDMETHOD(DeviceUnlock)();
        STDMETHOD(Enter)( VARIANT* Result,BSTR Format);
        STDMETHOD(Find)(BSTR Expression,long* Count,SAFEARRAY* Addresses);
        STDMETHOD( Initialize)(/*IInitialize*/void** pVal);//modified avoid IInitalize definition
        STDMETHOD(IOType)(BSTR* pVal);
        STDMETHOD(Output)(VARIANT OutputString);
        STDMETHOD(Read)(/*IRead*/void ** pVal);//avoid defining IRead
        STDMETHOD(BufferSize)(long* pVal);
        STDMETHOD(BufferSize)(long pVal);
 		STDMETHOD(ReadBytes)(long* Length,SAFEARRAY* Bytes);
        STDMETHOD(ReadTerminator)(short* pVal);
        STDMETHOD(ReadTerminator)( short pVal);
        STDMETHOD(Query)(VARIANT OutputString,VARIANT* ReturnVal);
        STDMETHOD(Timeout)(long* pVal);
        STDMETHOD(Timeout)(long pVal);
        STDMETHOD(Write)(/*IWrite*/void** pVal);//avoid IWrite definition
        STDMETHOD(WriteBytes)(long Length, SAFEARRAY* Data);
        STDMETHOD(WriteTerminator)(short* pVal);
        STDMETHOD(WriteTerminator)(short pVal);
    };


//This interface can be used with the 'FindSpecifiedInstrumentsIEnum' to get a
//collection of all instruments.  This example however does not show this as the
//more normal case is to know the address of the instrument you wish to 
//talk to.
interface IEnumIO : IUnknown {
        STDMETHOD(Next)( long celt, IIO** rgelt,long* pceltFetched);
        STDMETHOD(Skip)( long celt);
        STDMETHOD( Reset)();
        STDMETHOD(Clone)(IEnumIO** rgelt);
    };


//This defines the structure of the IIOManager interface (i.e. allows code such as this:
//IIOManager Iptr; ... Iptr->ConnectToInstrument(...) etc.)
interface IIOManager : IDispatch {
        
        STDMETHOD(ConnectToInstrument)(
												BSTR IOAddress, 
												IIO** ppIIO
									);
        STDMETHOD(FindSpecifiedInstruments)(
												BSTR Expression, 
												IUnknown** ppAgtIOServers
										);
        STDMETHOD(FindSpecifiedInstrumentsIEnum)(
												BSTR Expression, 
												IEnumIO** ppIEnumIO
												);
		
    };
}//extern "C"
/**************************************************************************************/


int main(int argc, char* argv[])
{
	IIOManager *IMngr;	//The IO Manager we'll be using to get to the instrument
	IIO *IDevice;		//The actual instrument.
	HRESULT hr;			//check return values for failure
	
	
	_bstr_t myInstrument,strTmp; //The string used to connect to the instrument
	_variant_t myCmd;
	_variant_t myResult;
	
	printf("Start CExample program.  We assume a 34401 voltmeter on COM1 set to 9600 baud\n\n");

	//Initialize COM
	CoInitialize(NULL);

	//create an instance of an AGTIOManager.  We wish to get the IIOManger interface
	//from this object.  The IO manager will allow us to get to any instrument
	//on the computer (assuming IntuiLink has been installed)
	hr = CoCreateInstance(CLSID_AGTIOManager,NULL,CLSCTX_INPROC_SERVER,IID_IIOManager,(LPVOID *)&IMngr);
	if FAILED(hr)
	{
		printf("CoCreateInstance failed\n");
		return 0;
	}

	IDevice=NULL;
	
	//this string asks for an instrument on RSR232 port 1, 9600 baud, dtr/dsr handshake
	myInstrument="COM1::BAUD=9600,PARITY=NONE,SIZE=8,HANDSHAKE=DTR_DSR";
	
	//Ask the IO manager for our particular instrument
	//a GPIB address would look like "GPIB::22" etc.

	hr = IMngr->ConnectToInstrument(myInstrument,&IDevice);
	if (FAILED(hr) || !IDevice)
	{
		IMngr->Release();
		printf("ConnectToInstrument failed\n");
		return 0;
	}


	//Its ok to release the IO manager were done with it
	IMngr->Release();

	//ok we have a device.  Now lets do an IDN? to be sure its a 34401

	myCmd="*IDN?";
	hr=IDevice->Query(myCmd,&myResult);//shows use of Query method
	if (FAILED(hr) )
	{
		IDevice->Release();
		printf("Query failed\n");
		return 0;
	}

	strTmp=myResult; //get the string into a bstr which can cast to c strings

	//now see if it contains a 34401, if so we assume we got the right instrument
	if (strstr(strTmp,"34401") == NULL )
	{
		IDevice->Release ();
		printf("Didn't find a 34401 voltmeter!\n");
		return 0;
	}

	//Before sending any real commands 34401 MUST be in remote mode
	hr = IDevice->Output(_variant_t("SYST:REMOTE"));
	if (FAILED(hr) )
	{
		IDevice->Release();
		printf("Output failed\n");
		return 0;
	}

	//set the instrument to dc volts and take a measurment
	hr = IDevice->Output(_variant_t("MEAS:VOLT:DC?"));
	if (FAILED(hr) )
	{
		IDevice->Release();
		printf("Output failed\n");
		return 0;
	}

	hr=IDevice->Enter(&myResult,_bstr_t("K"));  //'K' is the standard format
	if (FAILED(hr) )
	{
		IDevice->Release();
		printf("Enter failed\n");
		return 0;
	}

	//note the myResult string will have <cr><lf> at the end.
	printf("Voltage reading: %s",(char *)_bstr_t(myResult));

	//release the device before exit
	IDevice->Release();


	//Uninit so COM is exited clean.
	CoUninitialize();
	return 0;
}
