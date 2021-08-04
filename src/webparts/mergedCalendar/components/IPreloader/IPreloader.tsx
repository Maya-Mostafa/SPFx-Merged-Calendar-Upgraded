import * as React from 'react';
import styles from '../MergedCalendar.module.scss';
import { IPreloaderProps } from './IPreloaderProps';

import {Spinner, SpinnerSize, Overlay} from '@fluentui/react';

export default function IPreloader (props:IPreloaderProps) {

    return(
        <>
            {props.isDataLoading &&
                <>
                    <Overlay></Overlay>
                    <div className={styles.marginT20}>
                        <Spinner size={SpinnerSize.medium} label="Please Wait, Calendars are updating..." ariaLive="assertive" labelPosition="right" />
                    </div>
                </>
            }
        </>
    );
}