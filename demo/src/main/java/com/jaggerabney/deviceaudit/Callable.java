package com.jaggerabney.deviceaudit;

// simple SAM interface used in conjunction with functionWithProgressText to effectively pass a function as an argument 
public interface Callable<T> {
    public T call();
}
