import * as React from 'react';
import {
    MessageBar,
    MessageBarType
} from 'office-ui-fabric-react/lib/components/MessageBar';
import { DefaultButton } from 'office-ui-fabric-react/lib/components/Button';

export interface IAnnouncementProps {
    title: string;
    announcement: string;
    category: string;
    acknowledge: () => void;
}

export default class Announcement extends React.Component<IAnnouncementProps, {}> {
  constructor(props) {
    super(props);
  }

  public render(): JSX.Element {

    var messageBarType: MessageBarType;

    switch(this.props.category.toLowerCase())
    {
      case "blocked":
      {
        messageBarType = MessageBarType.blocked;
        break;
      }

      case "error":
      {
        messageBarType = MessageBarType.error;
        break;
      }

      case "info":
      {
        messageBarType = MessageBarType.info;
        break;
      }

      case "remove":
      {
        messageBarType = MessageBarType.remove;
        break;
      }

      case "success":
      {
        messageBarType = MessageBarType.success;
        break;
      }

      case "warning":
      {
        messageBarType = MessageBarType.warning;
        break;
      }

      default:
      {
        messageBarType = MessageBarType.success;
      }
    }

    return <MessageBar
            messageBarType={messageBarType}
            isMultiline={false}
            onDismiss={null}
            actions={<DefaultButton onClick={this.props.acknowledge}>OK</DefaultButton>}>
            <strong>{this.props.title}</strong>&nbsp;
            <span dangerouslySetInnerHTML={{__html: this.props.announcement.replace(/https?:[^\s]+/g, '<a href="$&">$&</a>')}} />
        </MessageBar>;
  }
}