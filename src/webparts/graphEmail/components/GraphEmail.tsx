import * as React from 'react';
import styles from './GraphEmail.module.scss';
import { IGraphEmailProps } from './IGraphEmailProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IMessages } from './IMessages';
import { IMessage } from './IMessage';

export default class GraphEmail extends React.Component<IGraphEmailProps, any> {
  constructor() {
    super();
    this.state = {
      messages: []
    }
  }

  private _loadMessages(): void {
    if (!this.props.graphClient) {
      return;
    }

    this.setState({
      error: null,
      loading: true,
      messages: []
    });
    this.props.graphClient
      .api("me/messages")
      .version("v1.0")
      .select("bodyPreview, receivedDateTime, from, subject, webLink")
      .top(5)
      .orderby("receivedDateTime desc")
      .get((err: any, res: IMessages): void => {
        if (err) {
          this.setState({
            error: err.messaage ? err.message: "Error", 
            loading: false
          });
          alert(err);
          return;
        }

        if (res && res.value && res.value.length > 0 ) {
          this.setState({
            messages: res.value,
            loading: false
          });
        }
        else 
        {
          this.setState({
            loading: false
          });
        }
      })
  }

  public componentDidMount(): void {
    this._loadMessages();
  }

  public render(): React.ReactElement<IGraphEmailProps> {
    var col = (this.state.messages)?this.state.messages:[]; 

    return (
      <div className={ styles.graphEmail }>
        <div className={ styles.container }>
         {col.map((item: IMessage) => {
                return (
  
                  <div className="ms-Grid-col ms-sm12 ms-md3">
                    item.from
                  </div>
                )
              })}
        </div>
      </div>
    );
  }
}
