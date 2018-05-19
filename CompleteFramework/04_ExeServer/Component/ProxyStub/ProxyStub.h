

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.00.0603 */
/* at Tue Mar 27 04:27:08 2018
 */
/* Compiler settings for C:\Users\Veni\Documents\Aashish\GIT\COM_Project\InvidualComponents\05_ExeServer\SERVER\ProxyStub.idl:
    Oicf, W1, Zp8, env=Win64 (32b run), target_arch=AMD64 8.00.0603 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __ProxyStub_h__
#define __ProxyStub_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IAddSubtract_FWD_DEFINED__
#define __IAddSubtract_FWD_DEFINED__
typedef interface IAddSubtract IAddSubtract;

#endif 	/* __IAddSubtract_FWD_DEFINED__ */


#ifndef __IMultiplyDivide_FWD_DEFINED__
#define __IMultiplyDivide_FWD_DEFINED__
typedef interface IMultiplyDivide IMultiplyDivide;

#endif 	/* __IMultiplyDivide_FWD_DEFINED__ */


/* header files for imported files */
#include "unknwn.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IAddSubtract_INTERFACE_DEFINED__
#define __IAddSubtract_INTERFACE_DEFINED__

/* interface IAddSubtract */
/* [unique][helpstring][uuid][object] */ 


EXTERN_C const IID IID_IAddSubtract;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("A3A50D32-58B3-4897-A447-DC8B935EFB94")
    IAddSubtract : public IUnknown
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE SumOfTwoIntegers( 
            /* [in] */ int __MIDL__IAddSubtract0000,
            /* [in] */ int __MIDL__IAddSubtract0001,
            /* [out][in] */ int *__MIDL__IAddSubtract0002) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IAddSubtractVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IAddSubtract * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IAddSubtract * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IAddSubtract * This);
        
        HRESULT ( STDMETHODCALLTYPE *SumOfTwoIntegers )( 
            IAddSubtract * This,
            /* [in] */ int __MIDL__IAddSubtract0000,
            /* [in] */ int __MIDL__IAddSubtract0001,
            /* [out][in] */ int *__MIDL__IAddSubtract0002);
        
        END_INTERFACE
    } IAddSubtractVtbl;

    interface IAddSubtract
    {
        CONST_VTBL struct IAddSubtractVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IAddSubtract_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IAddSubtract_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IAddSubtract_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IAddSubtract_SumOfTwoIntegers(This,__MIDL__IAddSubtract0000,__MIDL__IAddSubtract0001,__MIDL__IAddSubtract0002)	\
    ( (This)->lpVtbl -> SumOfTwoIntegers(This,__MIDL__IAddSubtract0000,__MIDL__IAddSubtract0001,__MIDL__IAddSubtract0002) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IAddSubtract_INTERFACE_DEFINED__ */


#ifndef __IMultiplyDivide_INTERFACE_DEFINED__
#define __IMultiplyDivide_INTERFACE_DEFINED__

/* interface IMultiplyDivide */
/* [unique][helpstring][uuid][object] */ 


EXTERN_C const IID IID_IMultiplyDivide;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("4D44D171-6898-4C54-BDEF-8302EA46ACBA")
    IMultiplyDivide : public IUnknown
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE SubtractinOfTwoIntegers( 
            /* [in] */ int __MIDL__IMultiplyDivide0000,
            /* [in] */ int __MIDL__IMultiplyDivide0001,
            /* [out][in] */ int *__MIDL__IMultiplyDivide0002) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IMultiplyDivideVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IMultiplyDivide * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IMultiplyDivide * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IMultiplyDivide * This);
        
        HRESULT ( STDMETHODCALLTYPE *SubtractinOfTwoIntegers )( 
            IMultiplyDivide * This,
            /* [in] */ int __MIDL__IMultiplyDivide0000,
            /* [in] */ int __MIDL__IMultiplyDivide0001,
            /* [out][in] */ int *__MIDL__IMultiplyDivide0002);
        
        END_INTERFACE
    } IMultiplyDivideVtbl;

    interface IMultiplyDivide
    {
        CONST_VTBL struct IMultiplyDivideVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IMultiplyDivide_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IMultiplyDivide_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IMultiplyDivide_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IMultiplyDivide_SubtractinOfTwoIntegers(This,__MIDL__IMultiplyDivide0000,__MIDL__IMultiplyDivide0001,__MIDL__IMultiplyDivide0002)	\
    ( (This)->lpVtbl -> SubtractinOfTwoIntegers(This,__MIDL__IMultiplyDivide0000,__MIDL__IMultiplyDivide0001,__MIDL__IMultiplyDivide0002) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IMultiplyDivide_INTERFACE_DEFINED__ */


/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


