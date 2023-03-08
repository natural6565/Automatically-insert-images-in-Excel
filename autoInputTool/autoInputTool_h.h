

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.01.0622 */
/* at Tue Jan 19 12:14:07 2038
 */
/* Compiler settings for autoInputTool.idl:
    Oicf, W1, Zp8, env=Win64 (32b run), target_arch=AMD64 8.01.0622 
    protocol : all , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */



/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 500
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif /* __RPCNDR_H_VERSION__ */


#ifndef __autoInputTool_h_h__
#define __autoInputTool_h_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IautoInputTool_FWD_DEFINED__
#define __IautoInputTool_FWD_DEFINED__
typedef interface IautoInputTool IautoInputTool;

#endif 	/* __IautoInputTool_FWD_DEFINED__ */


#ifndef __autoInputTool_FWD_DEFINED__
#define __autoInputTool_FWD_DEFINED__

#ifdef __cplusplus
typedef class autoInputTool autoInputTool;
#else
typedef struct autoInputTool autoInputTool;
#endif /* __cplusplus */

#endif 	/* __autoInputTool_FWD_DEFINED__ */


#ifdef __cplusplus
extern "C"{
#endif 



#ifndef __autoInputTool_LIBRARY_DEFINED__
#define __autoInputTool_LIBRARY_DEFINED__

/* library autoInputTool */
/* [version][uuid] */ 


EXTERN_C const IID LIBID_autoInputTool;

#ifndef __IautoInputTool_DISPINTERFACE_DEFINED__
#define __IautoInputTool_DISPINTERFACE_DEFINED__

/* dispinterface IautoInputTool */
/* [uuid] */ 


EXTERN_C const IID DIID_IautoInputTool;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("8aa0cea3-ccc4-4657-a525-3d2f6c01224b")
    IautoInputTool : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct IautoInputToolVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IautoInputTool * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IautoInputTool * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IautoInputTool * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IautoInputTool * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IautoInputTool * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IautoInputTool * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IautoInputTool * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } IautoInputToolVtbl;

    interface IautoInputTool
    {
        CONST_VTBL struct IautoInputToolVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IautoInputTool_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IautoInputTool_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IautoInputTool_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IautoInputTool_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IautoInputTool_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IautoInputTool_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IautoInputTool_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __IautoInputTool_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_autoInputTool;

#ifdef __cplusplus

class DECLSPEC_UUID("c4fb20f8-1b21-418e-8b79-5b251ec98ec9")
autoInputTool;
#endif
#endif /* __autoInputTool_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


