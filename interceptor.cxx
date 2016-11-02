#include "interceptor.hxx"
#include <com/sun/star/task/XJobListener.hpp>
#include <com/sun/star/ui/ContextMenuInterceptorAction.hpp>
#include <com/sun/star/ui/XContextMenuInterception.hpp>
#include <com/sun/star/ui/ContextMenuExecuteEvent.hpp>
#include <com/sun/star/ui/ActionTriggerSeparatorType.hpp>
#include <com/sun/star/beans/NamedValue.hpp>
#include <com/sun/star/frame/XFrame.hpp>
#include <com/sun/star/frame/XController.hpp>
#include <com/sun/star/frame/XModel.hpp>
#include <com/sun/star/uno/XComponentContext.hpp>
#include <com/sun/star/sheet/XSheetCellRanges.hpp>
#include <com/sun/star/sheet/XCellRangeAddressable.hpp>
#include <com/sun/star/sheet/XSpreadsheetDocument.hpp>
#include <com/sun/star/sheet/XSpreadsheets.hpp>
#include <com/sun/star/sheet/XSpreadsheet.hpp>
#include <com/sun/star/container/XEnumerationAccess.hpp>
#include <com/sun/star/container/XEnumeration.hpp>
#include <com/sun/star/container/XIndexAccess.hpp>
#include <com/sun/star/table/CellRangeAddress.hpp>
#include <com/sun/star/table/XCell.hpp>
#include <com/sun/star/table/CellContentType.hpp>
#include <com/sun/star/container/XIndexContainer.hpp>
#include <com/sun/star/lang/XMultiServiceFactory.hpp>
#include <com/sun/star/beans/XPropertySet.hpp>

#include <cppuhelper/supportsservice.hxx>
#include <rtl/ustring.hxx>

#include <cstdint>
#include <map>

using rtl::OUString;
using namespace com::sun::star::uno;
using namespace com::sun::star::frame;
using namespace com::sun::star::sheet;
using namespace com::sun::star::table;
using namespace com::sun::star::container;
using namespace com::sun::star::ui;
using com::sun::star::beans::NamedValue;
using com::sun::star::beans::XPropertySet;
using com::sun::star::task::XJobListener;
using com::sun::star::lang::IllegalArgumentException;
using com::sun::star::lang::XMultiServiceFactory;



// This is the service name an Add-On has to implement
#define SERVICE_NAME "com.sun.star.task.AsyncJob"

void IncrementMarkedCellValues( const Reference< XFrame > &rxFrame );
void IncrementCellRanges( Reference< XSheetCellRanges >& xRanges );
void IncrementCellRange( Reference< XCellRangeAddressable >& xRange, Reference< XModel >& xModel );

// Helper functions for the implementation of UNO component interfaces.
OUString RegisterInterceptorJobImpl_getImplementationName()
throw (RuntimeException)
{
    return OUString ( IMPLEMENTATION_NAME );
}

Sequence< OUString > SAL_CALL RegisterInterceptorJobImpl_getSupportedServiceNames()
throw (RuntimeException)
{
    Sequence < OUString > aRet(1);
    OUString* pArray = aRet.getArray();
    pArray[0] =  OUString ( SERVICE_NAME );
    return aRet;
}

Reference< XInterface > SAL_CALL RegisterInterceptorJobImpl_createInstance( const Reference< XComponentContext > & rContext)
    throw( Exception )
{
    return (cppu::OWeakObject*) new RegisterInterceptorJobImpl( rContext );
}

// Implementation of the recommended/mandatory interfaces of a UNO component.
// XServiceInfo
OUString SAL_CALL RegisterInterceptorJobImpl::getImplementationName()
    throw (RuntimeException)
{
    return RegisterInterceptorJobImpl_getImplementationName();
}

sal_Bool SAL_CALL RegisterInterceptorJobImpl::supportsService( const OUString& rServiceName )
    throw (RuntimeException)
{
    return cppu::supportsService(this, rServiceName);
}

Sequence< OUString > SAL_CALL RegisterInterceptorJobImpl::getSupportedServiceNames(  )
    throw (RuntimeException)
{
    return RegisterInterceptorJobImpl_getSupportedServiceNames();
}

// XAsyncJob method implementations

void SAL_CALL RegisterInterceptorJobImpl::executeAsync( const Sequence<NamedValue>& rArgs,
					  const Reference<XJobListener>& rxListener )
    throw(IllegalArgumentException, RuntimeException)
{
    printf("DEBUG>>> Called executeAsync() : this = %p\n", this); fflush(stdout);

    sal_Int32 nNumNVs = rArgs.getLength();
    Sequence<NamedValue> aEnvironment;
    for ( sal_Int32 nIdx = 0; nIdx < nNumNVs; ++nIdx )
    {
	if ( rArgs[nIdx].Name.equalsAscii("Environment") )
	{
	    rArgs[nIdx].Value >>= aEnvironment;
	    break;
	}
    }

    if ( aEnvironment.hasElements() )
    {
	sal_Int32 nNumEnvEntries = aEnvironment.getLength();
	Reference< XFrame > xFrame;
	OUString aEnvType, aEventName;
	for ( sal_Int32 nIdx = 0; nIdx < nNumEnvEntries; ++nIdx )
	{
	    if ( aEnvironment[nIdx].Name.equalsAscii("Frame") )
		aEnvironment[nIdx].Value >>= xFrame;
	    else if ( aEnvironment[nIdx].Name.equalsAscii("EnvType") )
		aEnvironment[nIdx].Value >>= aEnvType;
	    else if ( aEnvironment[nIdx].Name.equalsAscii("EventName") )
		aEnvironment[nIdx].Value >>= aEventName;
	}
	printf( "DEBUG>>> xFrame = %p\n", xFrame.get()); fflush(stdout);
	if ( aEnvType.equalsAscii("DISPATCH") && xFrame.is() )
	{
	    if (  aEventName.equalsAscii("onEnableInterceptClick1") )
	    {
		sal_Int32 nNumCalls = 0;
		Reference< XContextMenuInterceptor > xInterceptor = getInterceptor( xFrame, nNumCalls );
		if ( nNumCalls > 1 )
		{
		    printf("DEBUG>>> Interceptor is already enabled\n");
		    fflush(stdout);
		}
		else
		{
		    Reference< XController > xController = xFrame->getController();
		    if ( xController.is() )
		    {
			Reference< XContextMenuInterception > xContextMenuInterception( xController, UNO_QUERY );
			if ( xContextMenuInterception.is() )
			{
			    xContextMenuInterception->registerContextMenuInterceptor( xInterceptor );
			    printf("DEBUG>>> Registered ContextMenuInterceptorImpl to current frame controller.\n"); fflush(stdout);
			}
		    }
		}
	    } else if ( aEventName.equalsAscii("onIncrementClick") ) {
		printf("DEBUG>>> Got request from DISPATCH envType and event name onIncrementClick, going to call IncrementCellValue().\n"); fflush(stdout);
		IncrementMarkedCellValues( xFrame );
	    }
	}
    }

    Sequence<NamedValue> aReturn;
    rxListener->jobFinished( Reference<com::sun::star::task::XAsyncJob>(this), makeAny(aReturn));
    
}

Reference< XContextMenuInterceptor >& RegisterInterceptorJobImpl::getInterceptor( const Reference<XFrame>& xFrame, sal_Int32& rCalls )
{
    static Reference< XContextMenuInterceptor > xInterceptor;
    static std::map<std::uintptr_t, sal_Int32> aFrame2Calls;
    if ( !xInterceptor.is() )
	xInterceptor = Reference< XContextMenuInterceptor >( (cppu::OWeakObject*) new ContextMenuInterceptorImpl(), UNO_QUERY );

    std::uintptr_t nFramePtrVal = reinterpret_cast<std::uintptr_t>( xFrame.get() );
    auto aItr = aFrame2Calls.find( nFramePtrVal );
    if ( aItr != aFrame2Calls.end() )
	rCalls = ++(aItr->second);
    else
    {
	aFrame2Calls[ nFramePtrVal ] = 1;
	rCalls = 1;
    }
    return xInterceptor;
}


void logError(const char* pStr)
{
    printf(pStr);
    fflush(stdout);
}

void IncrementMarkedCellValues( const Reference< XFrame > &rxFrame )
{
    if ( !rxFrame.is() )
    {
	logError("DEBUG>>> IncrementCellValue : rxFrame is invalid.\n");
	return;
    }
    
    Reference< XController > xCtrl = rxFrame->getController();
    if ( !xCtrl.is() )
    {
	logError("DEBUG>>> IncrementCellValue : xCtrl is invalid.\n");
	return;
    }

    Reference< XModel > xModel = xCtrl->getModel();
    if ( !xModel.is() )
    {
	logError("DEBUG>>> IncrementCellValue : xModel is invalid.\n");
	return;
    }

    Reference< XSheetCellRanges > xRanges( xModel->getCurrentSelection(), UNO_QUERY );
    if ( xRanges.is() ) // Multimarked
    {
	IncrementCellRanges( xRanges );
    }
    else // Simple mark
    {
	Reference< XCellRangeAddressable > xRange( xModel->getCurrentSelection(), UNO_QUERY );
	if ( !xRange.is() )
	{
	    logError("DEBUG>>> IncrementCellValue : Both xRanges and xRange are invalid.\n");
	    return;
	}
	IncrementCellRange( xRange, xModel );
    }
    
}

void IncrementCellRange( Reference< XCellRangeAddressable >& xRange, Reference< XModel >& xModel )
{
    CellRangeAddress aRange = xRange->getRangeAddress();
    Reference< XSpreadsheetDocument > xSpreadsheetDocument( xModel, UNO_QUERY );
    if ( !xSpreadsheetDocument.is() )
	return;

    Reference< XIndexAccess > xSpreadsheets( xSpreadsheetDocument->getSheets(), UNO_QUERY );
    if ( !xSpreadsheets.is() )
    {
	logError("DEBUG>>> IncrementCellRange : Cannot get xSpreadsheets.\n");
	return;
    }

    try
    {
	Any aSheet = xSpreadsheets->getByIndex(aRange.Sheet);
	Reference< XSpreadsheet > xSpreadsheet( aSheet, UNO_QUERY );

	for ( sal_Int32 nCol = aRange.StartColumn; nCol <= aRange.EndColumn; ++nCol )
	{
	    for ( sal_Int32 nRow = aRange.StartRow; nRow <= aRange.EndRow; ++nRow )
	    {
		Reference< XCell > xCell = xSpreadsheet->getCellByPosition(nCol, nRow);
		if ( !xCell.is() )
		    logError("DEBUG>>> IncrementCellRange : xCell is invalid.\n");
		else
		{
		    if ( xCell->getType() == CellContentType_VALUE )
			xCell->setValue( xCell->getValue() + 1.0 );
		    else
			logError("DEBUG>>> IncrementCellRange : cell type is not numeric, skipping.\n");
		}
	    }
	}
    }
    catch( Exception& e )
    {
	fprintf(stderr, "DEBUG>>> IncrementCellRange : caught UNO exception: %s\n",
		OUStringToOString( e.Message, RTL_TEXTENCODING_ASCII_US ).getStr());
	fflush(stderr);
    }

    
}

void IncrementCellRanges( Reference< XSheetCellRanges >& xRanges )
{
    Reference< XEnumerationAccess > xEnumAccess = xRanges->getCells();
    if ( !xEnumAccess.is() )
    {
	logError("DEBUG>>> IncrementCellRanges : xEnumAccess is invalid.\n");
	return;
    }

    Reference< XEnumeration > xEnum = xEnumAccess->createEnumeration();
    if ( !xEnum.is() )
    {
	logError("DEBUG>>> IncrementCellRanges : xEnum is invalid.\n");
	return;
    }

    while ( xEnum->hasMoreElements() )
    {
	try
	{
	    Reference< XCell > xCell( xEnum->nextElement(), UNO_QUERY );
	    if ( !xCell.is() )
		logError("DEBUG>>> IncrementCellRanges : xCell is invalid.\n");
	    else
	    {
		if ( xCell->getType() == CellContentType_VALUE )
		    xCell->setValue( xCell->getValue() + 1.0 );
		else
		    logError("DEBUG>>> IncrementCellRanges : cell type is not numeric, skipping.\n");
	    }
	}
	catch ( Exception& e )
	{
	    fprintf(stderr, "DEBUG>>> IncrementCellRanges : caught UNO exception: %s\n",
		    OUStringToOString( e.Message, RTL_TEXTENCODING_ASCII_US ).getStr());
	    fflush(stderr);
	}
    }
}


ContextMenuInterceptorAction SAL_CALL ContextMenuInterceptorImpl::notifyContextMenuExecute(
    const ContextMenuExecuteEvent& rEvent )
    throw ( RuntimeException )
{

    printf("DEBUG>>> Inside notifyContextMenuExecute : this = %p\n", this); fflush(stdout);
    try {
	Reference< XIndexContainer > xContextMenu = rEvent.ActionTriggerContainer;
	if ( !xContextMenu.is() )
	{
	    logError("DEBUG>>> notifyContextMenuExecute : bad rEvent.ActionTriggerContainer\n");
	    return ContextMenuInterceptorAction_IGNORED;
	}
	Reference< XMultiServiceFactory > xMenuElementFactory( xContextMenu, UNO_QUERY );
	if ( !xMenuElementFactory.is() )
	{
	    logError("DEBUG>>> notifyContextMenuExecute : bad xMenuElementFactory\n");
	    return ContextMenuInterceptorAction_IGNORED;
	}

	Reference< XPropertySet > xSeparator( xMenuElementFactory->createInstance( "com.sun.star.ui.ActionTriggerSeparator" ), UNO_QUERY );
	if ( !xSeparator.is() )
	{
	    logError("DEBUG>>> notifyContextMenuExecute : cannot create xSeparator\n");
	    return ContextMenuInterceptorAction_IGNORED;
	}

	xSeparator->setPropertyValue( "SeparatorType", makeAny( ActionTriggerSeparatorType::LINE ) );
	
	Reference< XPropertySet > xMenuEntry( xMenuElementFactory->createInstance( "com.sun.star.ui.ActionTrigger" ), UNO_QUERY );
	if ( !xMenuEntry.is() )
	{
	    logError("DEBUG>>> notifyContextMenuExecute : cannot create xMenuEntry\n");
	    return ContextMenuInterceptorAction_IGNORED;
	}

	xMenuEntry->setPropertyValue( "Text", makeAny(OUString("Increment cell(s)")));
	xMenuEntry->setPropertyValue( "CommandURL", makeAny( OUString("vnd.sun.star.job:event=onIncrementClick") ) );

	sal_Int32 nCount = xContextMenu->getCount();
	sal_Int32 nIdx = nCount;
	xContextMenu->insertByIndex( nIdx++, makeAny( xSeparator ) );
	xContextMenu->insertByIndex( nIdx++, makeAny( xMenuEntry ) );

	return ContextMenuInterceptorAction_CONTINUE_MODIFIED;
    }
    catch ( Exception& e )
    {
	fprintf(stderr, "DEBUG>>> notifyContextMenuExecute : caught UNO exception: %s\n",
		OUStringToOString( e.Message, RTL_TEXTENCODING_ASCII_US ).getStr());
	fflush(stderr);
    }

    return ContextMenuInterceptorAction_IGNORED;
}

