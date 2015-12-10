#pragma once

// scope release helper
// not really a full-featured comptr

template <typename T>
class CComPtr
{
	T* m_ptr;
	CComPtr(CComPtr &other); // nope
	CComPtr& operator=(CComPtr &other); // nope

public:
	CComPtr() : m_ptr(nullptr) {}
	CComPtr(T *ptr) : m_ptr(ptr) {}
	~CComPtr()
	{
		ReleasePtr();
	}
	
	CComPtr(CComPtr &&other)
	{
		m_ptr = other.m_ptr;
		other.m_ptr = nullptr;
	}

	inline void ReleasePtr()
	{
		if (m_ptr)
		{
			m_ptr->Release();
			m_ptr = nullptr;
		}
	}

	inline operator bool() { return m_ptr != nullptr; }

	inline T* operator->() { return m_ptr; }

	inline T** operator&() { return &m_ptr; }

	inline T* Detach()
	{
		T* tmp = m_ptr;
		m_ptr = nullptr;
		return tmp;
	}
};
